/* global Office */
Office.onReady(() => {});

/** Ribbon button (Message Read): open a reply and then inject a draft */
export async function replyWithCopilot(event) {
  try {
    const item = Office.context.mailbox.item;
    await new Promise((resolve, reject) => {
      item.displayReplyFormAsync("", r => (r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error)));
    });
    // In compose, our OnNewMessageCompose event will fire and inject a draft
  } finally {
    event.completed();
  }
}

/** Ribbon button in Compose: regenerate */
export async function regenerateDraft(event) {
  try {
    await injectDraftFromService(true);
  } catch (e) {
    console.error(e);
  } finally {
    event.completed();
  }
}

/** Shared helpers used by events.js too (kept here to avoid CORS for function file) */
async function injectDraftFromService(regenerate) {
  const item = Office.context.mailbox.item;
  const subject = item.normalizedSubject || "";
  const bodyHtml = await getBodyHtml(item);

  // Load style profile (roaming settings)
  const profile = await loadStyleProfile();
  const draft = await generateDraftClient(subject, bodyHtml, profile, regenerate);

  // Prepend the draft above quoted thread
  await prependHtml(item, draft);
}

function getBodyHtml(item) {
  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Html, r => r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error));
  });
}
function prependHtml(item, html) {
  return new Promise((resolve, reject) => {
    item.body.prependAsync(html, { coercionType: Office.CoercionType.Html }, r => r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error));
  });
}

/** ---- Minimal client-side LLM wiring ----
 * If you paste an Azure OpenAI/OpenAI key in the Taskpane settings, we use it.
 * Otherwise, we fall back to a deterministic, non-LLM template generator.
 */
async function generateDraftClient(subject, replyHtml, profile, regenerate) {
  const settings = await getAddinSettings();
  const tl = profile?.prefs?.tldr !== false; // default true
  const tone = profile?.prefs?.tone || "direct";
  const length = profile?.prefs?.length || "medium";
  const closing = profile?.prefs?.closing || "Thanks";

  // Extract last human-written message heuristically
  const latest = stripQuoted(replyHtml);

  // If API key present, call LLM
  if (settings?.apiProvider && settings?.apiKey && settings?.endpoint) {
    try {
      const draftText = await callLLM(settings, {
        subject, latest,
        prefs: { tone, length, closing, tldr: tl }
      });
      return renderHtmlDraft(draftText);
    } catch (e) {
      console.warn("LLM call failed, using fallback:", e);
    }
  }

  // Fallback template (non-LLM)
  const summary = truncate(latest.replace(/<[^>]+>/g, ""), 350);
  const tldrBlock = tl ? `<p><strong>TL;DR:</strong> ${guessTldr(summary)}</p>` : "";
  const body = `${tldrBlock}
<p>Hi,</p>
<p>${generateBodyFromHeuristics(summary, tone, length)}</p>
<p>${closing},<br>${Office.context.mailbox.userProfile.displayName || ""}</p>`;
  return wrapDraft(body);
}

function stripQuoted(html) {
  // basic approach: cut at common markers
  const markers = [/From:\s/i, /Sent:\s/i, /Subject:\s/i, /<div class="?OutlookMessageHeader"?/i, /-----Original Message-----/i];
  let text = html;
  for (const m of markers) {
    const idx = text.search(m);
    if (idx > 0) { text = text.slice(0, idx); break; }
  }
  return text;
}
function truncate(s, n) { return s.length <= n ? s : s.slice(0, n) + "…"; }
function guessTldr(s) { return s.length > 180 ? s.slice(0, 180) + "…" : s; }
function generateBodyFromHeuristics(s, tone, length) {
  const ask = (/need(?:\s+your)?\s+(?:input|approval|decision)|by\s+\w+\s+\d{1,2}|blocked/i.exec(s) || [""])[0];
  const bullets = [`Address ${ask || "the ask"} or confirm next step`, `Owner & date`, `Call out any risk`]
    .map(x => `<li>${x}</li>`).join("");
  const opener = tone === "direct" ? "Thanks for the details — here’s my response:" : "Appreciate the context — here are my thoughts:";
  const lenTxt = length === "medium" ? "Keeping it concise:" : "Quickly:";
  return `${opener} ${lenTxt}<ul>${bullets}</ul>`;
}
function renderHtmlDraft(text) { return wrapDraft(text); }
function wrapDraft(innerHtml) {
  return `<div style="font-family:Segoe UI,Arial,sans-serif;font-size:12pt;line-height:1.4">${innerHtml}<hr/></div>`;
}

/** Simple settings + style profile using RoamingSettings (per mailbox) */
async function getAddinSettings() {
  return new Promise((resolve) => {
    const s = Office.context.roamingSettings.get("replycopilot.settings") || {};
    resolve(s);
  });
}
async function loadStyleProfile() {
  return new Promise((resolve) => {
    const profile = Office.context.roamingSettings.get("replycopilot.profile") || {
      prefs: { tone: "direct", length: "medium", closing: "Thanks", tldr: true },
      style: { bullets: 0.6, greeting: "Hi" }
    };
    resolve(profile);
  });
}
async function saveStyleProfile(profile) {
  return new Promise((resolve) => {
    Office.context.roamingSettings.set("replycopilot.profile", profile);
    Office.context.roamingSettings.saveAsync(() => resolve(true));
  });
}

/** Lightweight LLM call (supports Azure OpenAI or OpenAI) */
async function callLLM(settings, payload) {
  const sys = `You are a Senior PM. Tone: ${payload.prefs.tone}. Length: ${payload.prefs.length}. Closing: ${payload.prefs.closing}. Include TL;DR if requested. Be clear, action-oriented, and concise.`;
  const user = `Subject: ${payload.subject}\n\nLatest message:\n${payload.latest}\n\nWrite a reply that acknowledges, addresses asks/decisions, and lists next steps (owner/date).`;

  if (settings.apiProvider === "azure") {
    const url = `${settings.endpoint}openai/deployments/${settings.deployment}/chat/completions?api-version=${settings.apiVersion}`;
    const resp = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json", "api-key": settings.apiKey },
      body: JSON.stringify({ messages: [{ role: "system", content: sys }, { role: "user", content: user }], temperature: 0.3 })
    });
    const data = await resp.json();
    return data.choices?.[0]?.message?.content || "";
  } else {
    // OpenAI
    const resp = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${settings.apiKey}` },
      body: JSON.stringify({ model: settings.model || "gpt-4o-mini", messages: [{ role: "system", content: sys }, { role: "user", content: user }], temperature: 0.3 })
    });
    const data = await resp.json();
    return data.choices?.[0]?.message?.content || "";
  }
}

// Expose some helpers to events.js via global (if loaded separately)
window._replycopilot = { injectDraftFromService, saveStyleProfile };
``
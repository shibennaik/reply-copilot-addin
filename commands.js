/* global Office, window, fetch */
Office.onReady(() => {});

/** Ribbon command on message READ: opens Reply compose; draft will be injected by onNewMessageCompose */
async function replyWithCopilot(event) {
  try {
    const item = Office.context.mailbox.item;
    await new Promise((resolve, reject) => {
      item.displayReplyFormAsync("", r => (r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error)));
    });
  } catch (e) {
    console.error(e);
  } finally {
    event.completed();
  }
}

/** Ribbon command in COMPOSE: regenerate draft on demand */
async function regenerateDraft(event) {
  try {
    if (window._replycopilot?.injectDraftFromService) {
      await window._replycopilot.injectDraftFromService(true);
    }
  } catch (e) {
    console.error(e);
  } finally {
    event.completed();
  }
}

/* ---------------- Shared helpers (used by events.js too) ---------------- */

async function injectDraftFromService(regenerate) {
  const item = Office.context.mailbox.item;

  const subject = item.normalizedSubject || "";
  const replyHtml = await getBodyHtml(item);

  // Load style/profile prefs (with your defaults baked in)
  const profile = await loadStyleProfile();

  // Generate draft (client side; can call LLM if you add a key in Taskpane)
  const draftHtml = await generateDraftClient(subject, replyHtml, profile, regenerate);

  // Prepend the draft above the quoted thread
  await prependHtml(item, draftHtml);
}

function getBodyHtml(item) {
  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Html, r =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error)
    );
  });
}
function prependHtml(item, html) {
  return new Promise((resolve, reject) => {
    item.body.prependAsync(html, { coercionType: Office.CoercionType.Html }, r =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error)
    );
  });
}

/** ---- Draft generation (LLM optional; otherwise a clean heuristic template) ---- */
async function generateDraftClient(subject, replyHtml, profile, regenerate) {
  const settings = await getAddinSettings();

  // Preferences (defaults: your choices)
  const tone = profile?.prefs?.tone || "direct";
  const length = profile?.prefs?.length || "medium";
  const closing = profile?.prefs?.closing || "Thanks";
  const includeTldr = profile?.prefs?.tldr !== false;

  // Extract the latest human-written message from the reply chain
  const latestHtml = stripQuotedHtml(replyHtml);
  const latestText = htmlToText(latestHtml);

  // Try LLM if configured in Taskpane
  if (settings?.apiProvider && settings?.apiKey && (settings.endpoint || settings.apiProvider === "openai")) {
    try {
      const llmText = await callLLM(settings, {
        subject,
        latest: latestText,
        prefs: { tone, length, closing, tldr: includeTldr }
      });
      return wrapDraftHtml(llmText);
    } catch (e) {
      console.warn("LLM call failed; using local template instead:", e);
    }
  }

  // Local deterministic draft (no network)
  const tldr = includeTldr ? `<p><strong>TL;DR:</strong> ${guessTldr(latestText)}</p>` : "";
  const greeting = pickGreeting(profile);
  const body = composeHeuristicBody({ latestText, tone, length });
  const signature = buildSignature(profile);

  const html = `
    ${tldr}
    <p>${greeting},</p>
    ${body}
    <p>${closing},<br>${Office.context.mailbox.userProfile.displayName || ""}</p>
    ${signature ? `<div style="margin-top:8px">${signature}</div>` : ""}
  `;
  return wrapDraftHtml(html);
}

function stripQuotedHtml(html) {
  // Remove common quoted blocks; keep only the top (latest) portion
  const markers = [
    /<div[^>]*class=["']?OutlookMessageHeader["']?[^>]*>/i,
    /<hr[^>]*>/i,
    /-----Original Message-----/i,
    /<blockquote/i,
    /<div[^>]*id=["']?divRplyFwdMsg["']?[^>]*>/i,
    /<table[^>]*class=["']?MsoNormalTable["']?[^>]*>/i
  ];
  let cut = html;
  for (const m of markers) {
    const idx = cut.search(m);
    if (idx > 0) { cut = cut.slice(0, idx); break; }
  }
  return cut;
}
function htmlToText(html) {
  return (html || "")
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}
function guessTldr(text) {
  if (!text) return "Acknowledged. See response below.";
  const trimmed = text.length > 180 ? text.slice(0, 180) + "…" : text;
  return trimmed;
}
function composeHeuristicBody({ latestText, tone, length }) {
  const opener =
    tone === "direct"
      ? "Here’s my response and next steps:"
      : tone === "warm"
      ? "Thanks for the context—here are my thoughts and next steps:"
      : "Sharing a concise response and next steps:";

  const ask =
    (/need(?:\s+your)?\s+(?:input|approval|decision)/i.exec(latestText) ||
      /by\s+\w{3,9}\s+\d{1,2}/i.exec(latestText) ||
      /blocked/i.exec(latestText) || [""])[0];

  const bulletCount = length === "long" ? 4 : length === "concise" ? 2 : 3;
  const bullets = [
    ask ? `Address ${ask}` : "Confirm the ask and decision/answer",
    "List clear next steps with owners and dates",
    "Call out risks/assumptions",
    "Offer to sync if anything is unclear"
  ]
    .slice(0, bulletCount)
    .map(li => `<li>${li}</li>`)
    .join("");

  return `<p>${opener}</p><ul>${bullets}</ul>`;
}
function pickGreeting(profile) {
  return (profile?.style?.greeting || "Hi");
}
function buildSignature(profile) {
  return profile?.signatureHtml || "";
}
function wrapDraftHtml(inner) {
  return `<div style="font-family:Segoe UI,Arial,sans-serif;font-size:12pt;line-height:1.4">${inner}<hr/></div>`;
}

/* ---- Settings & style profile (stored locally via RoamingSettings) ---- */
async function getAddinSettings() {
  return new Promise((resolve) => {
    resolve(Office.context.roamingSettings.get("replycopilot.settings") || {});
  });
}
async function loadStyleProfile() {
  return new Promise((resolve) => {
    const def = {
      prefs: { tone: "direct", length: "medium", closing: "Thanks", tldr: true },
      style: { greeting: "Hi" },
      signatureHtml: ""
    };
    resolve(Office.context.roamingSettings.get("replycopilot.profile") || def);
  });
}
async function saveStyleProfile(profile) {
  return new Promise((resolve) => {
    Office.context.roamingSettings.set("replycopilot.profile", profile);
    Office.context.roamingSettings.saveAsync(() => resolve(true));
  });
}

/** Optional LLM call (Azure OpenAI or OpenAI) */
async function callLLM(settings, payload) {
  const system = `You are a Senior Product Manager. Tone: ${payload.prefs.tone}. Length: ${payload.prefs.length}. Closing: ${payload.prefs.closing}. Include a TL;DR if requested. Be concise, direct, and action-oriented.`;
  const user = `Subject: ${payload.subject}\n\nLatest message:\n${payload.latest}\n\nWrite a reply that:\n1) Acknowledges sender\n2) Addresses each ask/decision\n3) Lists next steps with owner/date\n4) Uses my tone and length preferences.`;

  if (settings.apiProvider === "azure") {
    const url = `${settings.endpoint.replace(/\/+$/,'/') }openai/deployments/${settings.deployment}/chat/completions?api-version=${settings.apiVersion}`;
    const resp = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json", "api-key": settings.apiKey },
      body: JSON.stringify({
        messages: [{ role: "system", content: system }, { role: "user", content: user }],
        temperature: 0.3
      })
    });
    const data = await resp.json();
    return data?.choices?.[0]?.message?.content || "";
  }

  if (settings.apiProvider === "openai") {
    const resp = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${settings.apiKey}` },
      body: JSON.stringify({
        model: settings.model || "gpt-4o-mini",
        messages: [{ role: "system", content: system }, { role: "user", content: user }],
        temperature: 0.3
      })
    });
    const data = await resp.json();
    return data?.choices?.[0]?.message?.content || "";
  }

  throw new Error("No provider configured");
}

/* Expose for events.js */
window._replycopilot = { injectDraftFromService, saveStyleProfile };

/* Expose ribbon commands globally (required by Outlook) */
window.replyWithCopilot = replyWithCopilot;
window.regenerateDraft = regenerateDraft;
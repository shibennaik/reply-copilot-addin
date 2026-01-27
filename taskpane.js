/* global Office */
Office.onReady(() => {
  init();
});

function init() {
  const settings = Office.context.roamingSettings.get("replycopilot.settings") || {};
  const profile = Office.context.roamingSettings.get("replycopilot.profile") || {
    prefs: { tone: "direct", length: "medium", closing: "Thanks", tldr: true },
    style: { greeting: "Hi" },
    signatureHtml: ""
  };

  setVal("length", profile.prefs.length || "medium");
  setVal("tone", profile.prefs.tone || "direct");
  setVal("closing", profile.prefs.closing || "Thanks");
  setChecked("tldr", profile.prefs.tldr !== false);

  setVal("apiProvider", settings.apiProvider || "");
  setVal("endpoint", settings.endpoint || "");
  setVal("deployment", settings.deployment || "");
  setVal("apiVersion", settings.apiVersion || "");
  setVal("apiKey", settings.apiKey || "");
  setVal("model", settings.model || "");

  document.getElementById("save").addEventListener("click", saveAll);
  document.getElementById("reset").addEventListener("click", resetProfile);
}

function saveAll() {
  const profile = Office.context.roamingSettings.get("replycopilot.profile") || { prefs: {}, style: {} };
  profile.prefs = {
    length: getVal("length"),
    tone: getVal("tone"),
    closing: getVal("closing"),
    tldr: getChecked("tldr")
  };
  Office.context.roamingSettings.set("replycopilot.profile", profile);

  const settings = {
    apiProvider: getVal("apiProvider"),
    endpoint: getVal("endpoint"),
    deployment: getVal("deployment"),
    apiVersion: getVal("apiVersion"),
    apiKey: getVal("apiKey"),
    model: getVal("model")
  };
  Office.context.roamingSettings.set("replycopilot.settings", settings);

  Office.context.roamingSettings.saveAsync(() => {
    showSaved();
  });
}

function resetProfile() {
  Office.context.roamingSettings.remove("replycopilot.profile");
  Office.context.roamingSettings.saveAsync(() => {
    showSaved("Profile reset ✓");
  });
}

function showSaved(text) {
  const el = document.getElementById("saved");
  el.textContent = text || "Saved ✓";
  el.style.display = "inline";
  setTimeout(() => (el.style.display = "none"), 2000);
}

function setVal(id, v){ const el=document.getElementById(id); if(el) el.value=v; }
function getVal(id){ const el=document.getElementById(id); return el?el.value:""; }
function setChecked(id, v){ const el=document.getElementById(id); if(el) el.checked=!!v; }
function getChecked(id){ const el=document.getElementById(id); return !!document.getElementById(id)?.checked; }
``
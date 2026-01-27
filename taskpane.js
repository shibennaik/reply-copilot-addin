/* global Office */
Office.onReady(() => {
  init();
});

function init() {
  const s = Office.context.roamingSettings.get("replycopilot.settings") || {};
  const p = (Office.context.roamingSettings.get("replycopilot.profile") || { prefs: {} }).prefs;

  setVal("length", p.length || "medium");
  setVal("tone", p.tone || "direct");
  setVal("closing", p.closing || "Thanks");
  setChecked("tldr", p.tldr !== false);

  setVal("apiProvider", s.apiProvider || "");
  setVal("endpoint", s.endpoint || "");
  setVal("deployment", s.deployment || "");
  setVal("apiVersion", s.apiVersion || "");
  setVal("apiKey", s.apiKey || "");
  setVal("model", s.model || "");

  document.getElementById("save").addEventListener("click", () => {
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
      document.getElementById("saved").style.display = "block";
      setTimeout(() => document.getElementById("saved").style.display = "none", 2000);
    });
  });
}

function setVal(id, v){ const el=document.getElementById(id); if(el) el.value=v; }
function getVal(id){ const el=document.getElementById(id); return el?el.value:""; }
function setChecked(id, v){ const el=document.getElementById(id); if(el) el.checked=v; }
function getChecked(id){ const el=document.getElementById(id); return !!document.getElementById(id)?.checked; }
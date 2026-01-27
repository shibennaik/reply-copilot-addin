/* global Office, window */
Office.onReady(() => {});

/** Event: fires on Reply/New/Forward compose */
async function onNewMessageCompose(event) {
  try {
    // Inject a first draft automatically
    if (window._replycopilot?.injectDraftFromService) {
      await window._replycopilot.injectDraftFromService(false);
    }
  } catch (e) {
    console.error(e);
  } finally {
    event.completed();
  }
}

/** Event: fires on Send — learn from your sent mail (client-side) */
async function onMessageSend(event) {
  try {
    const item = Office.context.mailbox.item;
    const body = await new Promise((resolve, reject) => {
      item.body.getAsync(Office.CoercionType.Text, r => r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error));
    });
    // Extract a few style signals and update profile (kept local in RoamingSettings)
    const profile = Office.context.roamingSettings.get("replycopilot.profile") || { prefs: {}, style: {} };
    profile.style = profile.style || {};
    profile.style.avgLen = (profile.style.avgLen || 0) * 0.8 + body.length * 0.2;
    profile.style.usesBullets = /(^|\n)\s*[-•]/m.test(body) ? 1 : (profile.style.usesBullets || 0) * 0.9;
    Office.context.roamingSettings.set("replycopilot.profile", profile);
    Office.context.roamingSettings.saveAsync(() => {
      event.completed({ allowEvent: true });
    });
  } catch (e) {
    console.error(e);
    event.completed({ allowEvent: true }); // never block sending
  }
}

// Expose handlers for manifest
window.onNewMessageCompose = onNewMessageCompose;
window.onMessageSend = onMessageSend;
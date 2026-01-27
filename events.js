/* global Office, window */
Office.onReady(() => {});

/** Fires when Reply/New/Forward compose opens */
async function onNewMessageCompose(event) {
  try {
    if (window._replycopilot?.injectDraftFromService) {
      await window._replycopilot.injectDraftFromService(false);
    }
  } catch (e) {
    console.error(e);
  } finally {
    // Always complete within ~5s
    event.completed();
  }
}

/** Fires on Send: collect simple style signals and update profile (local only) */
async function onMessageSend(event) {
  try {
    const item = Office.context.mailbox.item;

    const bodyText = await new Promise((resolve, reject) => {
      item.body.getAsync(Office.CoercionType.Text, r =>
        r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value || "") : reject(r.error)
      );
    });

    const profile = Office.context.roamingSettings.get("replycopilot.profile") || { prefs: {}, style: {} };
    profile.style = profile.style || {};
    // Exponential moving averages for a couple of signals
    profile.style.avgLen = (profile.style.avgLen || bodyText.length) * 0.8 + bodyText.length * 0.2;
    const hasBullets = /(^|\n)\s*[-•]/m.test(bodyText);
    profile.style.usesBullets = (profile.style.usesBullets || 0.4) * 0.8 + (hasBullets ? 1 : 0) * 0.2;

    Office.context.roamingSettings.set("replycopilot.profile", profile);
    Office.context.roamingSettings.saveAsync(() => {
      event.completed({ allowEvent: true }); // never block sending
    });
  } catch (e) {
    console.error(e);
    event.completed({ allowEvent: true });
  }
}

/* Expose globally for Outlook’s event dispatcher */
window.onNewMessageCompose = onNewMessageCompose;
window.onMessageSend = onMessageSend;
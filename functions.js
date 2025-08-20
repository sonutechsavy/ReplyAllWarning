/* global Office */
// Reply-All Guard: blocks or warns when recipients > threshold
const THRESHOLD = 3;

Office.onReady(() => {
  // no-op
});

function getRecipientsCount(item) {
  return new Promise((resolve, reject) => {
    let toCount = 0, ccCount = 0;
    const done = () => resolve(toCount + ccCount);

    item.to.getAsync((toRes) => {
      if (toRes.status !== Office.AsyncResultStatus.Succeeded) return reject(toRes.error);
      toCount = (toRes.value || []).length;
      item.cc.getAsync((ccRes) => {
        if (ccRes.status !== Office.AsyncResultStatus.Succeeded) return reject(ccRes.error);
        ccCount = (ccRes.value || []).length;
        done();
      });
    });
  });
}

// Entry point declared in manifest
function checkRecipientsOnSend(event) {
  const item = Office.context.mailbox.item;
  getRecipientsCount(item).then((count) => {
    if (count > THRESHOLD) {
      const msg = `You're about to email ${count} recipients. Are you sure you need Reply All?`;
      // Show error bar and soft-block send
      item.notificationMessages.addAsync("ReplyAllGuard", {
        type: "errorMessage",
        message: msg
      });
      event.completed({ allowEvent: false });
    } else {
      event.completed({ allowEvent: true });
    }
  }).catch((err) => {
    // Fail open if we can't read recipients
    console.error("Reply-All Guard error:", err);
    event.completed({ allowEvent: true });
  });
}

// Expose
if (typeof window !== "undefined") {
  window.checkRecipientsOnSend = checkRecipientsOnSend;
}

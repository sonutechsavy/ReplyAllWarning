/* global Office */

// To + Cc combined threshold
const THRESHOLD = 3;

Office.onReady(() => {
  // UI-less runtime; nothing to render
});

/**
 * Returns Promise<number> for (To + Cc) count in compose.
 */
function getRecipientsCount(item) {
  return new Promise((resolve, reject) => {
    let toCount = 0, ccCount = 0;

    item.to.getAsync(toRes => {
      if (toRes.status !== Office.AsyncResultStatus.Succeeded) {
        return reject(toRes.error);
      }
      toCount = (toRes.value || []).length;

      item.cc.getAsync(ccRes => {
        if (ccRes.status !== Office.AsyncResultStatus.Succeeded) {
          return reject(ccRes.error);
        }
        ccCount = (ccRes.value || []).length;

        resolve(toCount + ccCount);
      });
    });
  });
}

/**
 * Declared in the manifest as the OnMessageSend handler.
 * Decide whether to allow send based on recipient count.
 */
function checkRecipientsOnSend(event) {
  try {
    const item = Office.context.mailbox.item;

    getRecipientsCount(item)
      .then(count => {
        if (count > THRESHOLD) {
          // Show native error bar and block send
          const msg = `You're about to email ${count} recipients (limit ${THRESHOLD}). Reduce recipients to send.`;
          item.notificationMessages.addAsync("ReplyAllGuard", {
            type: "errorMessage",
            message: msg
          });/* global Office */

// To + Cc combined threshold
const THRESHOLD = 3;

Office.onReady(() => {
  // UI-less runtime; nothing to render
});

/**
 * Returns Promise<number> for (To + Cc) count in compose.
 */
function getRecipientsCount(item) {
  return new Promise((resolve, reject) => {
    let toCount = 0, ccCount = 0;

    item.to.getAsync(toRes => {
      if (toRes.status !== Office.AsyncResultStatus.Succeeded) {
        return reject(toRes.error);
      }
      toCount = (toRes.value || []).length;

      item.cc.getAsync(ccRes => {
        if (ccRes.status !== Office.AsyncResultStatus.Succeeded) {
          return reject(ccRes.error);
        }
        ccCount = (ccRes.value || []).length;

        resolve(toCount + ccCount);
      });
    });
  });
}

/**
 * Declared in the manifest as the OnMessageSend handler.
 * Decide whether to allow send based on recipient count.
 */
function checkRecipientsOnSend(event) {
  try {
    const item = Office.context.mailbox.item;

    getRecipientsCount(item)
      .then(count => {
        if (count > THRESHOLD) {
          // Show native error bar and block send
          const msg = `You're about to email ${count} recipients (limit ${THRESHOLD}). Reduce recipients to send.`;
          item.notificationMessages.addAsync("ReplyAllGuard", {
            type: "errorMessage",
            message: msg
          });
          event.completed({ allowEvent: false }); // block send
        } else {
          event.completed({ allowEvent: true }); // allow send
        }
      })
      .catch(err => {
        // Fail open to avoid unexpected permanent blocks if API fails
        console.error("Reply-All Guard error:", err);
        event.completed({ allowEvent: true });
      });

  } catch (e) {
    console.error("Reply-All Guard exception:", e);
    event.completed({ allowEvent: true });
  }
}

// Expose for the runtime to find
if (typeof window !== "undefined") {
  window.checkRecipientsOnSend = checkRecipientsOnSend;
}

          event.completed({ allowEvent: false }); // block send
        } else {
          event.completed({ allowEvent: true }); // allow send
        }
      })
      .catch(err => {
        // Fail open to avoid unexpected permanent blocks if API fails
        console.error("Reply-All Guard error:", err);
        event.completed({ allowEvent: true });
      });

  } catch (e) {
    console.error("Reply-All Guard exception:", e);
    event.completed({ allowEvent: true });
  }
}

// Expose for the runtime to find
if (typeof window !== "undefined") {
  window.checkRecipientsOnSend = checkRecipientsOnSend;
}

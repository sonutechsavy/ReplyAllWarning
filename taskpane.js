Office.onReady(() => {
  console.log("Reply All Warning Add-in is ready.");
});

function onReplyAll(event) {
  if (!Office.context.mailbox || !Office.context.mailbox.item) {
    console.error("Mailbox or item context is not available.");
    event.completed();
    return;
  }

  const item = Office.context.mailbox.item;
  const threshold = 5;

  const toCount = item.to && Array.isArray(item.to) ? item.to.length : 0;
  const ccCount = item.cc && Array.isArray(item.cc) ? item.cc.length : 0;
  const totalRecipients = toCount + ccCount;

  if (totalRecipients >= threshold) {
    Office.context.ui.displayDialogAsync(
      "https://sonutechsavy.github.io/ReplyAllWarning/warning.html",
      { height: 40, width: 30 },
      result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const dialog = result.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, args => {
            if (args.message === "proceed") {
              item.displayReplyAllFormAsync();
              dialog.close();
            } else {
              dialog.close();
            }
          });
        } else {
          console.error("Dialog failed to open: " + result.error.message);
        }
      }
    );
  } else {
    item.displayReplyAllFormAsync();
  }

  event.completed();
}

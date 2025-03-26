Office.onReady(() => {
  console.log("Reply All Warning Add-in is ready.");
});

function onReplyAll(event) {
  const item = Office.context.mailbox.item;

  const threshold = 5;

  const toCount = item.to ? item.to.length : 0;
  const ccCount = item.cc ? item.cc.length : 0;
  const totalRecipients = toCount + ccCount;

  if (totalRecipients >= threshold) {
    Office.context.ui.displayDialogAsync(
      "https://sonutechsavy.github.io/ReplyAllWarning/warning.html",
      { height: 40, width: 30 },
      result => {
        const dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, args => {
          if (args.message === "proceed") {
            item.displayReplyAllFormAsync();
            dialog.close();
          } else {
            dialog.close();
          }
        });
      }
    );
  } else {
    item.displayReplyAllFormAsync();
  }

  event.completed();
}

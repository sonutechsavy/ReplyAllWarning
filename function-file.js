Office.onReady(info => {
    if (info.host === Office.HostType.Outlook) {
      // Assign event handlers and perform any initial setup
    }
  });
  
  function warnReplyAll(event) {
    let item = Office.context.mailbox.item;
    let recipients = item.to.concat(item.cc); // Combine To and CC recipients
    if (recipients.length > 5) { // Example threshold
      let response = confirm('Are you sure you want to reply to all?');
      if (!response) {
        event.completed({ allowEvent: false });
      } else {
        event.completed({ allowEvent: true });
      }
    } else {
      event.completed({ allowEvent: true });
    }
  }
  
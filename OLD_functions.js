Office.onReady(function (info) {
    if (info.host === Office.HostType.Mailbox) {
      // Assign event handler when the add-in is ready
      document.getElementById("replyWithTemplateButton").onclick = replyWithTemplate;
    }
  });
  
  function replyWithTemplate() {
    // Implement logic to reply with a template
    // Use Office.js API to interact with the Outlook item
  }
  
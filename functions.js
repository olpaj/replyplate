Office.onReady(function (info) {
    if (info.host === Office.HostType.Mailbox) {
      // Assign event handler when the add-in is ready
      document.getElementById("replyWithTemplateButton").onclick = replyWithTemplate;
    }
  });
  
  function replyWithTemplate() {
    // Use Office.js API to get the current item
    var item = Office.context.mailbox.item;
  
    // Get the file input element
    var templateFileInput = document.getElementById("templateFileInput");
  
    // Check if a file is selected
    if (templateFileInput.files.length > 0) {
      // Read the content of the selected file
      var templateFile = templateFileInput.files[0];
      var reader = new FileReader();
  
      reader.onload = function (event) {
        // Set the email body with the template content
        var template = event.target.result;
        item.body.setAsync(template, { coercionType: Office.CoercionType.Text }, function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error("Error setting template:", asyncResult.error.message);
          } else {
            // Send the reply
            item.displayReplyForm({ htmlBody: template });
          }
        });
      };
  
      reader.readAsText(templateFile);
    } else {
      console.error("No template file selected.");
    }
  }
  
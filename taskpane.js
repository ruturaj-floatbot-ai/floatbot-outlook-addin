function formatLinks(event) {
  Office.context.mailbox.item.body.getAsync("html", function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      let body = asyncResult.value;

      // Replace "LEXI here" with a link
      body = body.replace(/LEXI here/g, '<a href="https://floatbot.ai/experience-zone/" target="_blank">LEXI here</a>');

      // Replace "Grab time here" with a link
      body = body.replace(/Grab time here/g, '<a href="https://meetings.hubspot.com/ruturaj-rana" target="_blank">Grab time here</a>');

      // Set the updated body back
      Office.context.mailbox.item.body.setAsync(body, { coercionType: "html" }, function (setResult) {
        if (setResult.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Body updated successfully.");
        } else {
          console.error("Failed to update body:", setResult.error.message);
        }

        // Tell Outlook we're done
        event.completed();
      });
    } else {
      console.error("Failed to get body:", asyncResult.error.message);
      event.completed();
    }
  });
}

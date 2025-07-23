Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    // Office is ready
  }
});

function formatLinks(event) {
  Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      let originalBody = asyncResult.value;

      // Replace "LEXI" with link to Floatbot demo
      let updatedBody = originalBody.replace(/(LEXI)(?![^<]*>|[^<>]*<\/a>)/g, '<a href="https://floatbot.ai/experience-zone/" target="_blank">LEXI</a>');

      // Replace "grab time" or "Grab time" with your HubSpot calendar link
      updatedBody = updatedBody.replace(/(grab time)(?![^<]*>|[^<>]*<\/a>)/gi, '<a href="https://meetings.hubspot.com/ruturaj-rana" target="_blank">$1</a>');

      Office.context.mailbox.item.body.setAsync(updatedBody, { coercionType: Office.CoercionType.Html }, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Body updated.");
        } else {
          console.error("Failed to update body:", asyncResult.error);
        }
      });
    } else {
      console.error("Failed to get body:", asyncResult.error);
    }
    event.completed();
  });
}

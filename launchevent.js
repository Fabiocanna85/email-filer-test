/*
 * Email Filer (test) - runs when you press Send.
 * If the email has not been through the filing panel, the send is held
 * and the alert offers to open the panel. If it has, the send goes through.
 */

function onMessageSendHandler(event) {
  Office.context.mailbox.item.loadCustomPropertiesAsync(function (result) {

    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      // Could not check. Hold the send; the user can still choose Send Anyway.
      event.completed({
        allowEvent: false,
        errorMessage: "The filing check could not run. Choose Send Anyway to send without filing.",
        cancelLabel: "Don't send"
      });
      return;
    }

    var confirmed = result.value.get("filingConfirmed");

    if (confirmed === "yes") {
      event.completed({ allowEvent: true });
    } else {
      event.completed({
        allowEvent: false,
        errorMessage: "This email has not been filed. Select Take action to choose a file name and location, or Send Anyway to send without filing.",
        cancelLabel: "Don't send",
        commandId: "msgComposeOpenPaneButton"
      });
    }
  });
}

// Required: links the name used in the manifest to the function above.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

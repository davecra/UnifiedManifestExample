/* global Office, self, window, global */
Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

function wordFunctionClick(event) {
  Office.context.document.setSelectedDataAsync("The Word ExecuteFunction works. Button ID=" + event.source.id);
  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}

function excelFunctionClick(event) {
  Office.context.document.setSelectedDataAsync("The Excel ExecuteFunction works. Button ID=" + event.source.id);
  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}

// /**
//  * Shows a notification when the add-in command is executed.
//  * @param event {Office.AddinCommands.Event}
//  */
// function action(event) {
//   const message = {
//     type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
//     message: "Performed action.",
//     icon: "Icon.80x80",
//     persistent: true
//   };

//   // Show a notification message
//   Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

//   // Be sure to indicate when the add-in command function is complete
//   event.completed();
// }

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// the add-in command functions need to be available in global scope
g.wordFunctionClick = wordFunctionClick;
g.excelFunctionClick = excelFunctionClick;

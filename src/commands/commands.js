/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

console.log('Will this be triggered');

Office.initialize = () => {
  console.log('-----------< Commands Office.initialize');

  Word.run(context => {

    console.log('Context initialized 123');
    console.log(context);

     // insert a paragraph at the end of the document.
     const paragraph = context.document.body.insertParagraph("Command initialize", Word.InsertLocation.end);

     // change the paragraph color to blue.
     paragraph.font.color = "red";

     context.sync();
  });
}

Office.onReady(() => {
  // If needed, Office.js is ready to be called

  console.log('-----------< Commands Ready');
  console.log('-----------< Commands Ready');

  Word.run(context => {

    console.log('Context initialized 123');
    console.log(context);

     // insert a paragraph at the end of the document.
     const paragraph = context.document.body.insertParagraph("Command onReady", Word.InsertLocation.end);

     // change the paragraph color to blue.
     paragraph.font.color = "red";

     context.sync();
  });
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  console.log('Executed Action');

  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();

  console.log('Event completed action');
}

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

// The add-in command functions need to be available in global scope
g.action = action;

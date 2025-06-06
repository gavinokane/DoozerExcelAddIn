/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event) {
  // Display a "Hello World" notification
  Office.context.mailbox.item.notificationMessages.addAsync("helloWorld", {
    type: "informationalMessage",
    message: "Hello World",
    icon: "icon-16",
    persistent: false
  });

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);

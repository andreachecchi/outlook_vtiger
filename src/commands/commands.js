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
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const item = Office.context.mailbox.item;

  // Ottieni il corpo del messaggio
  item.body.getAsync("text", (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const body = result.value;
      // Recupera altre proprietà
      const subject = item.subject || "Nessun oggetto";
      const from = item.from ? item.from.emailAddress : "Mittente sconosciuto";
      const to = item.to ? item.to.map((t) => t.emailAddress).join(", ") : "Destinatari sconosciuti";

      // Costruisci il messaggio
      const rawMessage = `
              Oggetto: ${subject}
              Mittente: ${from}
              Destinatari: ${to}
              
              Corpo:
              ${body}
          `;

      // Visualizza l'alert
      alert(rawMessage);
    } else {
      console.error("Errore nel recuperare il corpo:", result.error.message);
    }
  });

  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Messaggio inserito in VTiger",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);

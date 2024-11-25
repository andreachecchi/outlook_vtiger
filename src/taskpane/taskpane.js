/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  /**
   * Insert your Outlook code here
   */

  const item = Office.context.mailbox.item;
  let insertAt = document.getElementById("item-subject");
  let label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
  insertAt.appendChild(label);
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode(item.subject));
  insertAt.appendChild(document.createElement("br"));
}

export async function loadContent() {
  const url = 'https://jsonplaceholder.typicode.com/posts/4'; // URL del webservice pubblico

  try {
    const response = await fetch(url);

      if (!response.ok) {
          throw new Error(`Errore HTTP: ${response.status}`);
      }

      const data = await response.json(); // Supponendo che il servizio restituisca JSON

      // Scrive i dati nel <div id="prova">
      const provaDiv = document.getElementById('prova');
      provaDiv.innerHTML = `
          <h3>${data.title}</h3>
          <p>${data.body}</p>
      `;
  } catch (error) {
      console.error('Errore nella chiamata al web service:', error);
      const provaDiv = document.getElementById('prova');
      provaDiv.textContent = 'Errore nel caricamento dei dati.';
  }
}

// Chiama la funzione all'avvio o su un evento specifico
loadContent();

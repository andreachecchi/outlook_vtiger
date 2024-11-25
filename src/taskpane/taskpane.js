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

async function getChallenge(basicauth_user, basicauth_pass, vt_url, vt_user) {
  const url = vt_url + "?operation=getchallenge&username=" + vt_user;
  const username = basicauth_user;
  const password = basicauth_pass;

  // eslint-disable-next-line no-undef
  auth = null;

  if (username != null && password != null) {
    // eslint-disable-next-line no-undef
    auth = "Basic " + Buffer.from(username + ":" + password).toString("base64");
  }

  try {
    // eslint-disable-next-line no-undef
    response = null;
    // eslint-disable-next-line no-undef
    if (auth == null) {
      // eslint-disable-next-line no-undef
      response = await fetch(url, {
        method: "GET",
      });
    } else {
      // eslint-disable-next-line no-undef
      response = await fetch(url, {
        method: "GET",
        headers: {
          // eslint-disable-next-line no-undef
          Authorization: auth,
        },
      });
    }

    // eslint-disable-next-line no-undef
    if (!response.ok) {
      // eslint-disable-next-line no-undef
      throw new Error("Error: " + response.statusText);
    }

    // eslint-disable-next-line no-undef
    const data = await response.json();

    // Controlla se la risposta contiene il token
    if (data.success && data.result && data.result.token) {
      return data.result.token; // Restituisci il token
    } else {
      throw new Error("Token non trovato nella risposta.");
    }
  } catch (error) {
    // eslint-disable-next-line no-undef
    console.error("Errore:", error.message);
    throw error; // Rilancia l'errore per una gestione successiva
  }
}

export async function run() {
  // eslint-disable-next-line no-undef
  const CryptoJS = require("crypto-js");

  let basicauth_user = document.getElementById("basicauth_user").value;
  let basicauth_pass = document.getElementById("basicauth_pass").value;
  let vt_url = document.getElementById("vt_url").value;
  let vt_user = document.getElementById("vt_user").value;
  let vt_accesskey = document.getElementById("vt_accesskey").value;

  Office.context.roamingSettings.set("basicauth_user", basicauth_user);
  Office.context.roamingSettings.set("basicauth_pass", basicauth_pass);
  Office.context.roamingSettings.set("vt_url", vt_url);
  Office.context.roamingSettings.set("vt_user", vt_user);
  Office.context.roamingSettings.set("vt_accesskey", vt_accesskey);

  // Salvataggio dei dati
  Office.context.roamingSettings.saveAsync(function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      console.log("Settings saved successfully!");
    } else {
      console.log("Error saving settings: " + result.error.message);
    }
  });

  getChallenge(basicauth_user, basicauth_pass, vt_url, vt_user)
    .then((token) => {
      // eslint-disable-next-line no-undef
      console.log("Token received:", token);

      const session = CryptoJS.MD5(token + vt_accesskey).toString(CryptoJS.enc.Hex);
    })
    .catch((error) => {
      // eslint-disable-next-line no-undef
      console.error("Error getting token:", error.message);
    });

  const item = Office.context.mailbox.item;
  item.body.getAsync("text", (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const body = result.value;
      const subject = item.subject || "No subject";
      const from = item.from ? item.from.emailAddress : "No sender";
      const to = item.to ? item.to.map((t) => t.emailAddress).join(", ") : "No receivers";

      const rawMessage = `
              Subject: ${subject}
              Sender: ${from}
              Receivers: ${to}
              
              Body:
              ${body}
          `;

      //TODO tutto
    } else {
      //TODO show alert
    }
  });

  /*
  let insertAt = document.getElementById("item-subject");
  let label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
  insertAt.appendChild(label);
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode(item.subject));
  insertAt.appendChild(document.createElement("br"));
  */
}

export async function loadContent() {
  document.getElementById("basicauth_user").value(Office.context.roamingSettings.get("basicauth_user"));
  document.getElementById("basicauth_pass").value(Office.context.roamingSettings.get("basicauth_pass"));
  document.getElementById("vt_url").value(Office.context.roamingSettings.get("vt_url"));
  document.getElementById("vt_user").value(Office.context.roamingSettings.get("vt_user"));
  document.getElementById("vt_accesskey").value(Office.context.roamingSettings.get("vt_accesskey"));

  /*
  const url = "https://jsonplaceholder.typicode.com/posts/4"; // URL del webservice pubblico

  try {
    const response = await fetch(url);

    if (!response.ok) {
      throw new Error(`Errore HTTP: ${response.status}`);
    }

    const data = await response.json(); // Supponendo che il servizio restituisca JSON

    // Scrive i dati nel <div id="prova">
    const provaDiv = document.getElementById("prova");
    provaDiv.innerHTML = `
          <h3>${data.title}</h3>
          <p>${data.body}</p>
      `;
  } catch (error) {
    console.error("Errore nella chiamata al web service:", error);
    const provaDiv = document.getElementById("prova");
    provaDiv.textContent = "Errore nel caricamento dei dati.";
  }
  */
}

// Chiama la funzione all'avvio o su un evento specifico
loadContent();

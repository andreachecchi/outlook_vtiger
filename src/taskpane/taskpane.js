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

    loadContent();
  }
});

async function getChallenge(basicauth_user, basicauth_pass, vt_url, vt_user) {
  const url = vt_url + "?operation=getchallenge&username=" + vt_user;

  // eslint-disable-next-line no-undef
  let auth = "";

  if (basicauth_user != "" && basicauth_pass != "") {
    const authString = `${basicauth_user}:${basicauth_pass}`;
    const encodedAuth = btoa(authString);
    auth = `Basic ${encodedAuth}`;
  }

  try {
    // eslint-disable-next-line no-undef
    let response = "";
    // eslint-disable-next-line no-undef
    if (auth == "") {
      // eslint-disable-next-line no-undef
      response = await fetch(url, {
        method: "GET",
        headers: {
          // eslint-disable-next-line no-undef
        },
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

async function getSession(basicauth_user, basicauth_pass, vt_url, vt_user, vt_tok_accesskey) {
  const url = vt_url; // + "?operation=login&username=" + vt_user + "&accessKey=" + vt_tok_accesskey;
  console.log("url:", url);

  // eslint-disable-next-line no-undef
  let auth = "";

  if (basicauth_user != "" && basicauth_pass != "") {
    const authString = `${basicauth_user}:${basicauth_pass}`;
    const encodedAuth = btoa(authString);
    auth = `Basic ${encodedAuth}`;
  }

  const params = new URLSearchParams();
  params.append("operation", "login");
  params.append("username", vt_user);
  params.append("accessKey", vt_tok_accesskey);

  try {
    // eslint-disable-next-line no-undef
    let response = "";
    // eslint-disable-next-line no-undef
    if (auth == "") {
      // eslint-disable-next-line no-undef
      response = await fetch(url, {
        method: "POST",
        headers: {
          // eslint-disable-next-line no-undef
          "Content-Type": "application/x-www-form-urlencoded",
        },
        body: params,
      });
    } else {
      // eslint-disable-next-line no-undef
      response = await fetch(url, {
        method: "POST",
        headers: {
          // eslint-disable-next-line no-undef
          Authorization: auth,
          "Content-Type": "application/x-www-form-urlencoded",
        },
        body: params,
      });
    }

    // eslint-disable-next-line no-undef
    if (!response.ok) {
      // eslint-disable-next-line no-undef
      throw new Error("Error: " + response.statusText);
    }

    // eslint-disable-next-line no-undef
    const data = await response.json();

    return data.result.sessionName; // Restituisci il result
  } catch (error) {
    // eslint-disable-next-line no-undef
    console.error("Errore:", error.message);
    throw error; // Rilancia l'errore per una gestione successiva
  }
}

async function getUserId(basicauth_user, basicauth_pass, vt_url, vt_session, vt_user) {
  const url = vt_url;
  // eslint-disable-next-line no-undef
  console.log("url:", url);

  // eslint-disable-next-line no-undef
  let auth = "";

  if (basicauth_user != "" && basicauth_pass != "") {
    const authString = `${basicauth_user}:${basicauth_pass}`;
    // eslint-disable-next-line no-undef
    const encodedAuth = btoa(authString);
    auth = `Basic ${encodedAuth}`;
  }

  let query = "select id from Users where user_name = '" + vt_user + "';";

  console.log("query: " + query);

  // eslint-disable-next-line no-undef
  const params = new URLSearchParams();
  params.append("operation", "query");
  params.append("sessionName", vt_session);
  params.append("query", query);

  try {
    // eslint-disable-next-line no-undef
    let response = "";
    // eslint-disable-next-line no-undef
    if (auth == "") {
      // eslint-disable-next-line no-undef
      response = await fetch(url+"?"+params.toString(), {
        method: "GET",
        //headers: {
        // eslint-disable-next-line no-undef
        //"Content-Type": "application/x-www-form-urlencoded",
        //},
      });
    } else {
      // eslint-disable-next-line no-undef
      response = await fetch(url+"?"+params.toString(), {
        method: "GET",
        headers: {
          // eslint-disable-next-line no-undef
          Authorization: auth,
          //"Content-Type": "application/x-www-form-urlencoded",
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

    return data.result[0].id; // Restituisci il result
  } catch (error) {
    // eslint-disable-next-line no-undef
    console.error("Errore:", error.message);
    throw error; // Rilancia l'errore per una gestione successiva
  }
}

async function addProjectTask(basicauth_user, basicauth_pass, vt_url, vt_session, email_subject, email_content, prj_id, vt_user) {
  const url = vt_url;
  console.log("url:", url);

  // eslint-disable-next-line no-undef
  let auth = "";

  if (basicauth_user != "" && basicauth_pass != "") {
    const authString = `${basicauth_user}:${basicauth_pass}`;
    const encodedAuth = btoa(authString);
    auth = `Basic ${encodedAuth}`;
  }

  let today = new Date();
  let year = today.getFullYear();
  let month = String(today.getMonth() + 1).padStart(2, '0');
  let day = String(today.getDate()).padStart(2, '0');
  let formattedTodayDate = `${year}-${month}-${day}`;

  let user_id = await getUserId(basicauth_user, basicauth_pass, vt_url, vt_session, vt_user);

  let element = {
    projecttaskname: "EMAIL - " + email_subject,
    description: email_content,
    projectid: prj_id,
    assigned_user_id: user_id,
    projecttaskhours: 0,
    startdate: formattedTodayDate,
  };
  const encodedElement = encodeURIComponent(JSON.stringify(element));

  const params = new URLSearchParams();
  params.append("operation", "create");
  params.append("sessionName", vt_session);
  params.append("elementType", "ProjectTask");
  params.append("element", JSON.stringify(element));

  try {
    // eslint-disable-next-line no-undef
    let response = "";
    // eslint-disable-next-line no-undef
    if (auth == "") {
      // eslint-disable-next-line no-undef
      response = await fetch(url, {
        method: "POST",
        headers: {
          // eslint-disable-next-line no-undef
          "Content-Type": "application/x-www-form-urlencoded",
        },
        body: params,
      });
    } else {
      // eslint-disable-next-line no-undef
      response = await fetch(url, {
        method: "POST",
        headers: {
          // eslint-disable-next-line no-undef
          Authorization: auth,
          "Content-Type": "application/x-www-form-urlencoded",
        },
        body: params,
      });
    }

    // eslint-disable-next-line no-undef
    if (!response.ok) {
      // eslint-disable-next-line no-undef
      throw new Error("Error: " + response.statusText);
    }

    // eslint-disable-next-line no-undef
    const data = await response.json();

    return data.result.sessionName; // Restituisci il result
  } catch (error) {
    // eslint-disable-next-line no-undef
    console.error("Errore:", error.message);
    throw error; // Rilancia l'errore per una gestione successiva
  }
}

async function getProjects(basicauth_user, basicauth_pass, vt_url, vt_session) {
  const url = vt_url;
  console.log("url:", url);

  // eslint-disable-next-line no-undef
  let auth = "";

  if (basicauth_user != "" && basicauth_pass != "") {
    const authString = `${basicauth_user}:${basicauth_pass}`;
    const encodedAuth = btoa(authString);
    auth = `Basic ${encodedAuth}`;
  }

  let query = "select * from Project where projectstatus != 'archived' order by createdtime desc;";

  const params = new URLSearchParams();
  params.append("operation", "query");
  params.append("sessionName", vt_session);
  params.append("query", query);

  try {
    // eslint-disable-next-line no-undef
    let response = "";
    // eslint-disable-next-line no-undef
    if (auth == "") {
      // eslint-disable-next-line no-undef
      response = await fetch(url+"?"+params.toString(), {
        method: "GET",
        //headers: {
        // eslint-disable-next-line no-undef
        //"Content-Type": "application/x-www-form-urlencoded",
        //},
      });
    } else {
      // eslint-disable-next-line no-undef
      response = await fetch(url+"?"+params.toString(), {
        method: "GET",
        headers: {
          // eslint-disable-next-line no-undef
          Authorization: auth,
          //"Content-Type": "application/x-www-form-urlencoded",
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

    return data.result; // Restituisci il result
  } catch (error) {
    // eslint-disable-next-line no-undef
    console.error("Errore:", error.message);
    throw error; // Rilancia l'errore per una gestione successiva
  }
}

export async function run() {
  try {

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

    let prj_id = document.getElementById("projectSelect").value;

    const token = await getChallenge(basicauth_user, basicauth_pass, vt_url, vt_user);
    console.log("Token received: ", token);
    console.log("vt_tok_accesskey (concat): ", token + vt_accesskey);
    const vt_tok_accesskey = CryptoJS.MD5(token + vt_accesskey).toString(CryptoJS.enc.Hex);
    console.log("vt_tok_accesskey: ", vt_tok_accesskey);
    const vt_session = await getSession(basicauth_user, basicauth_pass, vt_url, vt_user, vt_tok_accesskey);
    console.log("Session name:", vt_session);  
  
    const item = Office.context.mailbox.item;
    let rawMessage = "";
    item.body.getAsync("text", (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const body = result.value;
        const subject = item.subject || "No subject";
        const from = item.from ? item.from.emailAddress : "No sender";
        const to = item.to ? item.to.map((t) => t.emailAddress).join(", ") : "No receivers";

        rawMessage = `
                Subject: ${subject}
                Sender: ${from}
                Receivers: ${to}
                
                Body:
                ${body}
            `;
        addProjectTask(basicauth_user, basicauth_pass, vt_url, vt_session, subject, rawMessage, prj_id, vt_user);
        let saved_banner = document.getElementById("saved_banner");
        let archive_btn = document.getElementById("archive_btn");
        archive_btn.style.display = "none";
        saved_banner.style.display = "block";
        // eslint-disable-next-line no-undef
        setTimeout(function () {
          saved_banner.style.display = "none";
          archive_btn.style.display = "block";
        }, 3000);
      } else {
        alert("Generic error");
      }
    });
  } catch (error) {
    console.error("Error:", error.message);
  }
}

export async function loadContent() {
  // eslint-disable-next-line no-undef
  const CryptoJS = require("crypto-js");

  document.getElementById("basicauth_user").value = Office.context.roamingSettings.get("basicauth_user");
  document.getElementById("basicauth_pass").value = Office.context.roamingSettings.get("basicauth_pass");
  document.getElementById("vt_url").value = Office.context.roamingSettings.get("vt_url");
  document.getElementById("vt_user").value = Office.context.roamingSettings.get("vt_user");
  document.getElementById("vt_accesskey").value = Office.context.roamingSettings.get("vt_accesskey");

  let basicauth_user = document.getElementById("basicauth_user").value;
  let basicauth_pass = document.getElementById("basicauth_pass").value;
  let vt_url = document.getElementById("vt_url").value;
  let vt_user = document.getElementById("vt_user").value;
  let vt_accesskey = document.getElementById("vt_accesskey").value;

  const token = await getChallenge(basicauth_user, basicauth_pass, vt_url, vt_user);
  console.log("Token received: ", token);
  console.log("vt_tok_accesskey (concat): ", token + vt_accesskey);
  const vt_tok_accesskey = CryptoJS.MD5(token + vt_accesskey).toString(CryptoJS.enc.Hex);
  console.log("vt_tok_accesskey: ", vt_tok_accesskey);
  const vt_session = await getSession(basicauth_user, basicauth_pass, vt_url, vt_user, vt_tok_accesskey);
  console.log("Session name:", vt_session);

  const projects = await getProjects(basicauth_user, basicauth_pass, vt_url, vt_session);
  console.log(projects);

  let selectHTML = `<select id="projectSelect">`;
  projects.forEach(project => {
    selectHTML += `<option value="${project.id}">${project.projectname}</option>`;
  });
  selectHTML += `</select>`;
  document.getElementById("selectProjectContainer").innerHTML = selectHTML;
}

const config_btn = document.getElementById("config_btn");
const config_div = document.querySelector('.collapsible-content');
config_btn.addEventListener('click', () => {  
  config_div.classList.toggle('open');
});

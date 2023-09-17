import axios from "axios";
import { Buffer } from "node:buffer";
import { writeFileSync, unlinkSync, existsSync } from "node:fs";
import { exec } from "node:child_process";
import config from "./Config.js";

export async function GetLatestWowInformation(graphToken) {

  const response = await axios.get("https://graph.microsoft.com/v1.0/me/messages?$search=WOW", { headers: { Authorization: `Bearer ${graphToken}`}});

  let unfilteredWows = response.data.value;
  let filteredWows = unfilteredWows.filter(email => (email.subject.includes("WOW") && email.sender.emailAddress.address == config.from));
  console.log(JSON.stringify(filteredWows));
  return filteredWows[0];
  
}

/**
 * 
 * @param {string} graphToken Token to use with the Microsoft Graph API. 
 * @param {*} messageId The ID of the message (email) to get attachments for.
 * @returns {object[]} Array of attachments for the message.
 */
export async function GetMessageAttachments(graphToken, messageId) {

  const response = await axios.get(`https://graph.microsoft.com/v1.0/me/messages/${messageId}/attachments`, { headers: { Authorization: `Bearer ${graphToken}`}});
  return response;
}


export async function DownloadWow(graphToken, wowId) {

  await axios.get(`https://graph.microsoft.com/v1.0/me/messages/${wowId}/attachments`, { headers: { Authorization: `Bearer ${graphToken}`}}).then(response => {

    if (!response.data) {
      throw new Error("responsed.data was not present in the Axios response.");
    }

    console.log("Response.data: " + JSON.stringify(response.data));

    if (!response.data.value[0]) {
      throw new Error("response.data is present in the Axios but the array is empty. There are no attachments on this message. The WOW is probably linked in the message HTML.");
    }

    if (!response.data.value[0].contentBytes) {
      throw new Error("An attachment is present but it has no contentBytes.");
    }

    let buffer = Buffer.from(response.data.value[0].contentBytes, "base64");

    if (existsSync("wow.pptx")) {
      unlinkSync("wow.pptx");
    }

    writeFileSync("wow.pptx", buffer);
    exec("start powerpnt.exe /S wow.pptx");

  }).catch(e => {
    throw new Error(e);
  });
}
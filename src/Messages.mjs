import axios from "axios";
import { Buffer } from "node:buffer";
import { writeFileSync, unlinkSync, existsSync } from "node:fs";
import { exec } from "node:child_process";
import config from "./Config.js";

export async function GetLatestWowInformation(graphToken) {
  return new Promise(async (resolve, reject) => {

    await axios.get("https://graph.microsoft.com/v1.0/me/messages?$search=WOW", { headers: { Authorization: `Bearer ${graphToken}`}}).then(response => {
      let matchDate = new Date();
      let matchString = `WOW ${matchDate.getMonth()}/${matchDate.getDate()}`;

       let unfilteredWows = response.data.value;
       let filteredWows = unfilteredWows.filter(email => (email.subject.includes("WOW") && email.sender.emailAddress.address == config.from));
      
      resolve(filteredWows[0])
      
    });

  });
}

export async function GetMessageAttachments(graphToken, messageId) {
  return new Promise(async (resolve, reject) => {

    const response = await axios.get(`https://graph.microsoft.com/v1.0/me/messages/${messageId}/attachments`, { headers: { Authorization: `Bearer ${graphToken}`}});
    resolve(response.data.value);

  })
}

export async function DownloadWow(graphToken, wowId) {
  return new Promise(async (resolve, reject) => {
    await axios.get(`https://graph.microsoft.com/v1.0/me/messages/${wowId}/attachments`, { headers: { Authorization: `Bearer ${graphToken}`}}).then(response => {

      if (!response.data) {
        reject("responsed.data was not present in the Axios response.");
        return;
      }

      console.log("Response.data: " + JSON.stringify(response.data));

      if (!response.data.value[0]) {
        reject("response.data is present in the Axios but the array is empty. There are no attachments on this message. The WOW is probably linked in the message HTML.");
        return;
      }

      if (!response.data.value[0].contentBytes) {
        reject("An attachment is present but it has no contentBytes.")
        return;
      }

      let buffer = Buffer.from(response.data.value[0].contentBytes, 'base64');

      if (existsSync("wow.pptx")) {
        unlinkSync("wow.pptx");
      }

      writeFileSync("wow.pptx", buffer);
      exec("start powerpnt.exe /S wow.pptx");

    }).catch(e => {
      reject(e);
    });
  });
}
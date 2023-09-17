import axios from "axios";
import { createWriteStream } from 'fs';
import config from "../config.mjs";

export async function GetLatestSharedWow(graphToken, senderEmail) {
  return new Promise(async (resolve, reject) => {

    // Get recent shared items from OneDrive
    const response = await axios.get("https://graph.microsoft.com/v1.0/me/drive/sharedwithme", { headers: { Authorization: `Bearer ${graphToken}`}});

    // WOWs will always follow a "WOW [int]-[int].pptx" format.

    // Filter only files that are from designated email and match the name format.
    var sharedWows = response.data.value.filter(item => {
      return (item.createdBy.user.email == config.from && config.fileNameRegex.test(item.name));
    });

    // Check for none
    if (sharedWows.length == 0) {
      reject("No WOWs to show. It's probably some sort of holiday.")
    }
    
    resolve(sharedWows[0]); 

  });
}

// Used to get the unique one-time download URL for a drive item.
export async function GetDownloadUrl(graphToken, driveId, fileId) {
  return new Promise(async (resolve, reject) => {
    await axios.get(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}?select=id,@microsoft.graph.downloadUrl`, { headers: { Authorization: `Bearer ${graphToken}`}}).then(response => {
      resolve(response.data["@microsoft.graph.downloadUrl"]);
    });
  });
}

export async function DownloadFileFromUrl(url, filename) {
  return new Promise(async (resolve, reject) => {

    const response = await axios.get(url, { responseType: "stream"});
    response.data.pipe(createWriteStream(filename));

    response.data.on("end", () => { resolve(true); });
    response.data.on("error", () => { reject("Error while piping response to file."); });

  });
}
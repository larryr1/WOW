import axios from "axios";
import { createWriteStream } from "fs";
import config from "./Config.js";

export async function GetLatestSharedWow(graphToken) {

  // Get recent shared items from OneDrive
  const response = await axios.get("https://graph.microsoft.com/v1.0/me/drive/sharedwithme", { headers: { Authorization: `Bearer ${graphToken}`}});

  // WOWs will always follow a "WOW [int]-[int].pptx" format.

  // Filter only files that are from designated email and match the name format.
  let wowRegex = /WOW\s+[0-9]+-[0-9]+\.pptx/i;

  var sharedWows = response.data.value.filter(item => {
    return (item.createdBy.user.email == config.from && wowRegex.test(item.name));
  });

  // Check for none
  if (sharedWows.length == 0) {
    throw new Error("Could not locate any WOW emails.");
  }
    
  return sharedWows[0];
}

// Used to get the unique one-time download URL for a drive item.
export async function GetDownloadUrl(graphToken, driveId, fileId) {
  const response = await axios.get(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}?select=id,@microsoft.graph.downloadUrl`, { headers: { Authorization: `Bearer ${graphToken}`}});
  return response.data["@microsoft.graph.downloadUrl"];
}

export async function DownloadFileFromUrl(url, filename) {

  const response = await axios.get(url, { responseType: "stream"});
  return new Promise((resolve, reject) => {
    
    response.data.pipe(createWriteStream(filename));

    response.data.on("end", () => { resolve(true); });
    response.data.on("error", () => { reject("Error while piping response to file."); });
  });
  
  
}
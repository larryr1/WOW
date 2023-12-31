#!/usr/bin/env node
import config from "./Config.js";
import { GetAuthorizationCode, GetGraphToken } from "./Authorization.mjs";
import { GetLatestWowInformation, GetMessageAttachments } from "./Messages.mjs";
import { GetLatestSharedWow, GetDownloadUrl, DownloadFileFromUrl } from "./OneDrive.mjs";
import { RunPowerpoint, RunTransformer } from "./Transformer.js";
import { existsSync, unlinkSync, writeFileSync } from "fs";
import { fileTypeFromFile } from "file-type";


function authorizationCodeError(e) {
  throw new Error(e);
}

function graphTokenError(e) {
  throw new Error(e);
}

function getLatestWowInformationError(e) {
  throw new Error(e);
}


async function GetWowFromShared(graphToken) {

  console.log("Getting latest shared WOW.");
  const latestSharedFile = await GetLatestSharedWow(graphToken);

  console.log("Getting download url for shared WOW.");
  const fileDownloadUrl = await GetDownloadUrl(
    graphToken,
    latestSharedFile.remoteItem.parentReference.driveId,
    latestSharedFile.remoteItem.id
  );

  console.log("Downloading shared WOW.");
  await DownloadFileFromUrl(fileDownloadUrl, "wow.pptx");

}

(async () => {

  // Delete leftover files.
  console.log("Deleting old files.");
  if (existsSync("wow.pptx")) { unlinkSync("wow.pptx"); }

  // Obtain Microsoft authorization code.
  console.log("Obtaining Graph authorization code.");
  const authorizationCode = await GetAuthorizationCode().catch(authorizationCodeError);

  // Exchange authorization code for Microsoft Graph API access token.
  console.log("Exchanging authorization code for Graph access token.");
  const graphToken = await GetGraphToken(authorizationCode, config.clientSecret).catch(graphTokenError);

  // Query the SCPASub inbox to get the latest WOW email.
  console.log("Finding latest WOW.");
  const latestWowInformation = await GetLatestWowInformation(graphToken).catch(getLatestWowInformationError);

  console.log(`Latest WOW is from email "${latestWowInformation.subject}"`);

  // Take appropriate action depending on how the WOW was sent. Both routes download the wow to "wow.pptx".
  if (latestWowInformation.hasAttachments) {

    // Email has attachments and the WOW needs to be downloaded from the attachments.
    console.log("Downloading email attachment.");
    let attachments = await GetMessageAttachments(graphToken, latestWowInformation.id);
    console.log("Writing bytes.");
    console.log(attachments[0].name);
    let buffer = Buffer.from(attachments[0].contentBytes, "base64");

    if (existsSync("wow.pptx")) {
      unlinkSync("wow.pptx");
    }

    writeFileSync("wow.pptx", buffer);

  } else {

    // There is no attachment in the email and the WOW needs to be downloaded from the "Shared With Me" files.
    await GetWowFromShared(graphToken);

  }

  // Check to make sure downloaded file is a PowerPoint. The transformer throws cryptic errors if it's passed a non-pptx file.
  const fileType = await fileTypeFromFile("wow.pptx");
  const requiredMimeType = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  const requiredExtension = "pptx";

  try {
    if (fileType.mime != requiredMimeType) {
      throw new Error(`File MIME type ${fileType.mime} does not match required MIME type of ${requiredMimeType}.`);
    } else if (fileType.ext.toLowerCase() != requiredExtension) {
      throw new Error(`File extension .${fileType.ext} does not match required extension .${requiredExtension}.`);
    }
  } catch (error) {
    console.log("Error: " + error);
    
    // Check for an existing transformed WOW and use it
    if (existsSync("wow.pptx-transformed.pptx")) {
      console.log("Starting a cached WOW.");
      await RunPowerpoint("wow.pptx-transformed.pptx");
    } else {
      console.log("No WOW is cached. Aborting.");
    }

    return;
  }
  
  // Remove file before transformer
  if (existsSync("wow.pptx-transformed.pptx")) { unlinkSync("wow.pptx-transformed.pptx"); }

  // The transformer uses PowerPoint Interop DLLs to apply an automatic transition to every slide and sets the slideshow to loop.
  console.log("Running transformer.");
  await RunTransformer("wow.pptx");

  console.log("Waiting for transformer to close file.");  
  setTimeout(async () => {
    // Starts the PowerPoint in slideshow mode.
    console.log("Starting PowerPoint.");
    await RunPowerpoint("wow.pptx-transformed.pptx");
    
  }, 2000);

})();
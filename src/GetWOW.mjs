import config from "../config";
import { GetAuthorizationCode, GetGraphToken } from "./Authorization.mjs";
import { GetLatestWowInformation, GetMessageAttachments } from "./Messages.mjs";
import { GetLatestSharedWow, GetDownloadUrl, DownloadFileFromUrl } from "./OneDrive.mjs";
import { RunPowerpoint, RunTransformer } from "./Transformer.mjs";
import { existsSync, unlinkSync, writeFileSync } from 'fs';

const clientSecret = config.clientSecret;

function authorizationCodeError(e) {
  throw new Error(e);
}

function graphTokenError(e) {
  throw new Error(e);
}

function getLatestWowInformationError(e) {
  throw new Error(e);
}

function downloadWowError(e) {
  throw new Error(e);
}

async function GetWowFromShared(graphToken) {

  const latestSharedFile = await GetLatestSharedWow(graphToken);

  const fileDownloadUrl = await GetDownloadUrl(
    graphToken,
    latestSharedFile.remoteItem.parentReference.driveId,
    latestSharedFile.remoteItem.id
  );

  await DownloadFileFromUrl(fileDownloadUrl, "wow.pptx");

}

(async () => {

  // Delete leftover files.
  if (existsSync("wow.pptx")) { unlinkSync("wow.pptx"); }
  if (existsSync("wow.pptx-transformed.pptx")) { unlinkSync("wow.pptx-transformed.pptx"); }

  // Obtain Microsoft authorization code.
  const authorizationCode = await GetAuthorizationCode().catch(authorizationCodeError);

  // Exchange authorization code for Microsoft Graph API access token.
  const graphToken = await GetGraphToken(authorizationCode, config.clientSecret).catch(graphTokenError);

  // Query the SCPASub inbox to get the latest WOW email.
  const latestWowInformation = await GetLatestWowInformation(graphToken).catch(getLatestWowInformationError);

  // Take appropriate action depending on how the WOW was sent. Both routes download the wow to "wow.pptx".
  if (latestWowInformation.hasAttachments) {

    // Email has attachments and the WOW needs to be downloaded from the attachments.
    console.log("Get attachments.");
    let attachments = await GetMessageAttachments(graphToken, latestWowInformation.id);
    console.log("Writing content bytes for 1st attachment.");
    console.log(attachments[0].name)
    let buffer = Buffer.from(attachments[0].contentBytes, 'base64');

      if (existsSync("wow.pptx")) {
        unlinkSync("wow.pptx");
      }

      writeFileSync("wow.pptx", buffer);

    console.log("Downloaded raw attachment.");

  } else {

    // There is no attachment in the email and the WOW needs to be downloaded from the "Shared With Me" files.
    await GetWowFromShared(graphToken);

  }

  // The transformer uses PowerPoint Interop DLLs to apply an automatic transition to every slide and sets the slideshow to loop.
  console.log("Running transformer...");
  await RunTransformer("wow.pptx");

  console.log("Waiting for Transformer to close file.");  
  setTimeout(async () => {
    // Starts the PowerPoint in slideshow mode.
    console.log("Starting PowerPoint.");
    await RunPowerpoint("wow.pptx-transformed.pptx");
    
  }, 2000);

  

})();

export async function GetWow(options) {

  if (!options.fileName) {
    throw new Error("fileName is not present in the options argument.");
  }

  if (!options.client) {
    throw new Error("client object is not present in the options argument.");
  }

  if (!options.client.clientId) {
    throw new Error("client.clientId is not present in the options argument.");
  }

  if (!options.client.clientSecret) {
    throw new Error("client.clientSecret is not present in the options argument.");
  }
}
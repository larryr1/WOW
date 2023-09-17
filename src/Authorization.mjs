import puppeteer from "puppeteer";
import { parse } from "node:querystring";
import axios from "axios";
import { URLSearchParams } from "node:url";
import config from "./Config.js";

export async function GetAuthorizationCode() {

  // I can't figure out a way to refactor this to get rid of the promise here. I'll do it later.
  // eslint-disable-next-line no-async-promise-executor
  return new Promise(async (resolve, reject) => {

    // Puppeteer start
    const browser = await puppeteer.launch({ headless: !config.showPuppeteerWindow });
    const page = await browser.newPage();

    // Handler to capture the redirect from authorization
    page.setRequestInterception(true);

    page.on("request", async request => {

      let url = request.url();
      if (url.startsWith("http://localhost:24587")) {

        // Parse the querystring from the question mark forward
        var query = url.substring(url.indexOf("?") + 1);
        var accessCode = parse(query).code;
        await browser.close();

        resolve(accessCode);

      } else {

        // Ignore request if it is not our callback url.
        request.continue();

      }
    });

    try {

      // Authorization URL to obtain new token
      // This is specifically worked to navigate the login page of my organization. You may change it for your own needs.
      await page.goto(`https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/authorize?client_id=${config.clientId}&scope=Mail.read%20Files.Read.All&response_type=code&response_mode=query&login_hint=${config.email}`);
      // Fill in password
      const passwordInput = await page.waitForSelector("input#passwordInput");
      await passwordInput.type(config.password);
      await passwordInput.dispose();

      // Submit password
      const passwordSubmit = await page.waitForSelector("span#submitButton");
      await passwordSubmit.click();
      await passwordSubmit.dispose();

      // Accept application authorization prompt
      const declineStayButton = await page.waitForSelector("input[type='submit']");
      await declineStayButton.click();
      await declineStayButton.dispose(); 

    } catch (error) {
      reject(error);
    }

    // From here, the redirect handler will capture the redirect request.
    // That will be parsed and the code will be returned in the resolved promise.
    
  });

}


export async function GetGraphToken(authorizationCode, clientSecret) {

  // Ensure arguments
  if (!authorizationCode) { throw new Error("authorizationCode was not passed."); }
  if (!clientSecret) { throw new Error("clientSecret was not passed."); }

  // Obtain graph access token
  var accessParams = new URLSearchParams();
  accessParams.append("grant_type", "authorization_code");
  accessParams.append("client_id", config.clientId);
  accessParams.append("client_secret", config.clientSecret);
  accessParams.append("code", authorizationCode);

  // Request for graph token
  try {
    const response = await axios.post(`https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`, accessParams);

    // Ensure success
    if (!response.data) { throw new Error("A request error was uncaught, and response.data is not present."); }
    if (!response.data.access_token) { throw new Error("A request error was uncaught and response.data.access_token is not present."); }

    return response.data.access_token;

  } catch (error) {

    if (error.response) {
      console.log(`Response data: ${JSON.stringify(error?.response?.data)}`);
    }

    throw new Error("Error obtaining Graph access token: " + error);
  }
}
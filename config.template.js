export default config = {
  // The email address to expect new slides to come from.
  from: "",

  // A Regular Expression that dictates the expected file name of the slideshow.
  fileNameRegex: /WOW\s+[0-9]+-[0-9]+\.pptx/i,

  // Several options related to the Azure Auth application.
  // Information from the Azure application.
  tenantId: "",
  clientId: "",
  clientSecret: "",

  powerpointPath: "",

  // The login information for the account that is sent the slideshows
  email: "",
  password: "",

  // Obtaining authorization codes involves a hidden, controlled Chrome browser that simulates a real user. You may opt to show
  //    this browser window for better insight or debugging. Recommended to keep this false in production.
  showPuppeteerWindow: false
}
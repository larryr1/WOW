const path = require("path");

const configPath = path.join(process.cwd(), "./wow_config.json");

module.exports = require(configPath);
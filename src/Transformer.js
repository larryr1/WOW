const { exec } = require("child_process");
const config = require("./Config.js");
const path = require("path"); 

module.exports = {};

const RunPowerpoint = async (file) => {

  return new Promise((resolve, reject) => {
    exec(`"${config.powerpointPath}" /S "${file}"`, (error, stdout, stderr) => {
      if (error) {
        reject(error);
      }
      
      resolve(stdout? stdout : stderr);
    });
  });

};

const RunTransformer = async (file) => {
  return new Promise((resolve, reject) => {
    let transformerPath = path.resolve(__dirname, "PowerPointTransformer.exe") + " " + file;
    exec(transformerPath, (error, stdout, stderr) => {
      if (error) {
        reject(error);
      }
      
      resolve(stdout? stdout : stderr);
    });
  });
};

module.exports.RunTransformer = RunTransformer;
module.exports.RunPowerpoint = RunPowerpoint;
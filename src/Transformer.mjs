import { exec } from 'child_process';
import config from '../config.mjs';
import path from 'path';

export async function RunPowerpoint(file) {

  return new Promise((resolve, reject) => {
    exec(`"${config.powerpointPath}" /S "${file}"`, (error, stdout, stderr) => {
      if (error) {
        reject(error);
      }
      
      resolve(stdout? stdout : stderr);
    });
  });
}

export async function RunTransformer(file) {
  
  return new Promise((resolve, reject) => {
    let transformerPath = path.resolve("./src/PowerPointTransformer.exe") + " " + file;
    exec(transformerPath, (error, stdout, stderr) => {
      if (error) {
        reject(error);
      }
      
      resolve(stdout? stdout : stderr);
    });
  });
}
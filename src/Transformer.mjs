import { exec } from 'child_process';
import config from '../config';

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
    exec(`PowerPointTransformer.exe "${file}"`, (error, stdout, stderr) => {
      if (error) {
        reject(error);
      }
      
      resolve(stdout? stdout : stderr);
    });
  });
}
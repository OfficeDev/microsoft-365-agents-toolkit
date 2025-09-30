// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { setPassword } from "keytar";
import { EOL, networkInterfaces } from "os";
import { exit } from "process";
import { join } from "path";
import { readdir, stat } from "fs-extra";
import axios from "axios";

const testResourcesPath = join(__dirname, "..", "..", ".test-resources");

export async function getMachineId(): Promise<string> {
  let machineId: string;
  try {
    // find machineID
    const storage = require(join(
      testResourcesPath,
      "settings",
      "storage.json"
    ));
    machineId = storage["telemetry.machineId"];
  } catch (err) {
    const id = await getMacMachineId();
    machineId = id || "";
  }
  return machineId;
}

async function getMacMachineId(): Promise<string | undefined> {
  try {
    const crypto = await import("crypto");
    const macAddress = getMac();
    return crypto.createHash("sha256").update(macAddress, "utf8").digest("hex");
  } catch (err) {
    return undefined;
  }
}

const invalidMacAddresses = new Set([
  "00:00:00:00:00:00",
  "ff:ff:ff:ff:ff:ff",
  "ac:de:48:00:11:22",
]);

function validateMacAddress(candidate: string): boolean {
  const tempCandidate = candidate.replace(/\-/g, ":").toLowerCase();
  return !invalidMacAddresses.has(tempCandidate);
}

export function getMac(): string {
  const ifaces = networkInterfaces();
  for (const name in ifaces) {
    const networkInterface = ifaces[name];
    if (networkInterface) {
      for (const { mac } of networkInterface) {
        if (validateMacAddress(mac)) {
          return mac;
        }
      }
    }
  }

  throw new Error("Unable to retrieve mac address (unexpected format)");
}

const serviceName = "vscodems-vscode.azure-account";
const accountName = "AzureCloud";
const extensionId = "ms-vscode.azure-account";

const scopes = [
  "https://management.core.windows.net//user_impersonation",
  "email",
  "offline_access",
  "openid",
  "profile",
];
const clientID = "aebc6443-996d-45c2-90f0-388ff96faa56";
const grantType = "password";
//const tenentID = "72f988bf-86f1-41af-91ab-2d7cd011db47";
const tenentID = "989179d1-c44d-4cf0-bb69-e911c27db18a";
const baseURL = "https://login.microsoftonline.com";

async function findFileRecursively(path: string): Promise<string[]> {
  const result: string[] = [];
  const files = await readdir(path);
  for (const file of files) {
    const filePath = join(path, file);
    const fileStat = await stat(filePath);
    if (fileStat.isDirectory()) {
      result.push(...(await findFileRecursively(filePath)));
    } else {
      result.push(filePath);
    }
  }
  return result;
}

async function azureLogin() {
  // Entry
  const username = process.argv[3];
  const password = process.argv[4];

  if (!username || !password) {
    throw new Error(
      `Please provide username and password, e.g.,${EOL}\t npx ts-node azureLogin.ts -- "username" "password"`
    );
  }

  const client = axios.create({
    baseURL: baseURL,
    timeout: 1000 * 100,
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
  });

  const data = {
    client_id: clientID,
    scope: scopes.reduce((p: string, c: string, i: number) => {
      if (i == 0) {
        p = c;
      } else {
        p += ` ${c}`;
      }
      return p;
    }),
    username: username,
    password: password,
    grant_type: grantType,
  };

  const encodeForm = (data: any) => {
    return Object.keys(data)
      .map(
        (key) => encodeURIComponent(key) + "=" + encodeURIComponent(data[key])
      )
      .join("&");
  };

  const resp = await client.post(
    `/${tenentID}/oauth2/v2.0/token`,
    encodeForm(data)
  );

  const toEncrypt = JSON.stringify({
    extensionId,
    content: resp.data["refresh_token"],
  });

  const files = await findFileRecursively(testResourcesPath);
  let encryptFilePath: string;
  for (const file of files) {
    if (file.includes("vscode-encrypt-native.node")) {
      encryptFilePath = file;
    }
  }
  const machineId = await getMachineId();
  const encryption = require(encryptFilePath!);
  const encryptedData = await encryption.encrypt(machineId, toEncrypt);
  await setPassword(serviceName, accountName, encryptedData);
  console.log("Azure login Successfully!");
}

azureLogin().catch((err) => {
  console.error(err);
  exit(-1);
});

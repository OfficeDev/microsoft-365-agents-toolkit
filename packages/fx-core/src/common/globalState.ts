// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as os from "os";
import * as path from "path";
import * as fs from "fs-extra";
import crypto from "crypto";
import { ConfigFolderName, ProductName } from "@microsoft/teamsfx-api";
import properLock from "proper-lockfile";
import { waitSeconds } from "./utils";

const GlobalStateFileName = "state.json";

/**
 * Return a value.
 *
 * @param key A string.
 * @return The stored value or `undefined`.
 */
export async function globalStateGet(key: string, defaultValue?: any): Promise<any> {
  const filePath = getGlobalStateFile();
  ensureGlobalStateFileExists(filePath);

  const lockFileDir = getLockFolder(filePath);
  const lockfilePath = path.join(lockFileDir, `${ConfigFolderName}.lock`);
  await fs.ensureDir(lockFileDir);

  const retryNum = 10;
  for (let i = 0; i < retryNum; ++i) {
    try {
      await properLock.lock(filePath, { lockfilePath: lockfilePath });
      let value: any = undefined;
      try {
        const config = await fs.readJSON(filePath);
        value = config[key];
        if (value === undefined) {
          value = defaultValue;
        }
      } finally {
        await properLock.unlock(filePath, { lockfilePath: lockfilePath });
      }
      return value;
    } catch (e: any) {
      if (e["code"] === "ELOCKED") {
        await waitSeconds(1);
        continue;
      }
      return undefined;
    }
  }
}

/**
 * Store a value. The value must be JSON-stringifyable.
 *
 * @param key A string.
 * @param value A value. MUST not contain cyclic references.
 */
export async function globalStateUpdate(key: string, value: any): Promise<void> {
  const filePath = getGlobalStateFile();
  ensureGlobalStateFileExists(filePath);

  const lockFileDir = getLockFolder(filePath);
  const lockfilePath = path.join(lockFileDir, `${ConfigFolderName}.lock`);
  await fs.ensureDir(lockFileDir);

  const retryNum = 10;
  for (let i = 0; i < retryNum; ++i) {
    try {
      await properLock.lock(filePath, { lockfilePath: lockfilePath });
      try {
        const config = await fs.readJSON(filePath);
        config[key] = value;
        await fs.writeJson(filePath, config);
      } finally {
        await properLock.unlock(filePath, { lockfilePath: lockfilePath });
      }
      break;
    } catch (e: any) {
      if (e["code"] === "ELOCKED") {
        await waitSeconds(1);
        continue;
      }
      throw e;
    }
  }
}

function getGlobalStateFile(): string {
  const homeDir = os.homedir();
  return path.join(homeDir, `.${ConfigFolderName}`, GlobalStateFileName);
}

function ensureGlobalStateFileExists(filePath: string): void {
  if (!fs.pathExistsSync(path.dirname(filePath))) {
    fs.mkdirpSync(path.dirname(filePath));
  }

  if (!fs.existsSync(filePath)) {
    fs.writeJSONSync(filePath, {});
  }
}

function getLockFolder(projectPath: string): string {
  return path.join(
    os.tmpdir(),
    `${ProductName}-${crypto.createHash("sha256").update(projectPath).digest("hex")}`
  );
}

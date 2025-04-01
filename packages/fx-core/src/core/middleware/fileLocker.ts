// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";
import * as fs from "fs-extra";
import { lock } from "proper-lockfile";
import { getLockFolder } from "./concurrentLocker";
import path from "path";
import { ConfigFolderName } from "@microsoft/teamsfx-api";
import { waitSeconds } from "../../common/utils";

export async function withFileLock<T>(filePath: string, callback: () => Promise<T>): Promise<T> {
  if (!(await fs.pathExists(filePath))) {
    throw new Error(`File not found: ${filePath}`);
  }

  const lockFileDir = getLockFolder(filePath);
  const lockfilePath = path.join(lockFileDir, `${ConfigFolderName}.lock`);
  await fs.ensureDir(lockFileDir);

  let release: (() => Promise<void>) | null = null;
  for (let i = 0; i < 10; i++) {
    try {
      release = await lock(filePath, { lockfilePath: lockfilePath });
      break;
    } catch (e) {
      if (e.code === "ELOCKED") {
        await waitSeconds(1);
      } else {
        throw e;
      }
    }
  }

  if (!release) {
    throw new Error(`Failed to acquire lock on ${filePath} after 10 seconds.`);
  }

  try {
    return await callback();
  } finally {
    await release();
  }
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";
import * as fs from "fs-extra";
import { lock } from "proper-lockfile";

export async function withFileLock<T>(filePath: string, callback: () => Promise<T>): Promise<T> {
  if (!(await fs.pathExists(filePath))) {
    throw new Error(`File not found: ${filePath}`);
  }

  let release: (() => Promise<void>) | null = null;

  for (let i = 0; i < 10; i++) {
    try {
      release = await lock(filePath);
      break;
    } catch (e) {
      if (e["code"] === "ELOCKED") {
        await new Promise((resolve) => setTimeout(resolve, 1000));
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

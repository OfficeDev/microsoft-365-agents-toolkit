// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Xiaofu Huang <xiaofhua@microsoft.com>
 */
import * as fs from "fs-extra";
import { v3DefaultHelpLink } from "../constant/helpLink";
import { Messages } from "../constant/message";
import { DepsCheckerError } from "../../../error/depCheck";

export async function unlinkSymlink(linkFilePath: string, forceUpdate = false): Promise<void> {
  try {
    const stat = await fs.lstat(linkFilePath);
    if (stat.isSymbolicLink()) {
      await fs.unlink(linkFilePath);
    } else {
      if (forceUpdate) {
        // For regular file or directory, remove it
        await fs.remove(linkFilePath);
      }
    }
  } catch (error: unknown) {
    const statError = error as { code?: string };
    if (statError.code !== "ENOENT") {
      throw error;
    }
  }
}

export async function createSymlink(
  target: string,
  linkFilePath: string,
  forceUpdate = false
): Promise<void> {
  await unlinkSymlink(linkFilePath, forceUpdate);
  // check if destination already exists
  if (await fs.pathExists(linkFilePath)) {
    throw new DepsCheckerError(Messages.symlinkDirAlreadyExist(linkFilePath), v3DefaultHelpLink);
  }

  return await fs.ensureSymlink(
    target,
    linkFilePath,
    /* Only used for Windows. Directory junction is similar to directory link but does not require admin permission. */
    "junction"
  );
}

export async function rename(oldPath: string, newPath: string): Promise<void> {
  if (await fs.pathExists(newPath)) {
    await fs.remove(newPath);
  }
  await fs.rename(oldPath, newPath);
}

// sliently remove file or dir
export async function cleanup(path: string): Promise<void> {
  try {
    await fs.remove(path);
  } catch {}
}

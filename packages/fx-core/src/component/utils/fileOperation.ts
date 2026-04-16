// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs-extra";
import klaw from "klaw";
import AdmZip, { EntryHeader } from "adm-zip";
import { Ignore } from "ignore";
import path from "path";
import { CacheFileInUse, DeployEmptyFolderError, ZipFileError } from "../../error/deploy";

/**
 * Asynchronously zip a folder and return buffer
 * @param sourceDir base dir
 * @param notIncluded block list
 * @param cache zip cache file location
 */
export async function zipFolderAsync(
  sourceDir: string,
  cache: string,
  notIncluded: Ignore
): Promise<fs.ReadStream> {
  const tasks: Promise<void>[] = [];
  const ig = notIncluded;
  // always delete cache if exists
  if (fs.existsSync(cache)) {
    try {
      await fs.remove(cache);
    } catch (e) {
      if (e instanceof Error && (e as any)?.code === "EBUSY") {
        throw new CacheFileInUse(cache, e);
      }
      throw e;
    }
  }
  const zip = new AdmZip();

  const addFileIntoZip = async (
    zp: AdmZip,
    filePath: string,
    zipPath: string,
    stats?: fs.Stats
  ) => {
    const content = await fs.readFile(filePath);
    zp.addFile(zipPath, content);
    if (stats) {
      const entry = zp.getEntry(zipPath);
      if (entry) {
        (entry.header as EntryHeader).time = stats.mtime;
      }
    }
  };

  await forEachFileAndDir(
    sourceDir,
    (itemPath: string, stats: fs.Stats) => {
      const relativePath: string = path.relative(sourceDir, itemPath);
      const zipPath = path.normalize(relativePath).split("\\").join("/");
      if (relativePath && !stats.isDirectory()) {
        const fullPath = path.join(sourceDir, relativePath);
        const task = addFileIntoZip(zip, fullPath, zipPath, stats);
        tasks.push(task);
      }
    },
    (itemPath: string) => {
      return !ig.test(path.relative(sourceDir, itemPath)).ignored;
    }
  );

  if (tasks.length === 0) {
    throw new DeployEmptyFolderError(sourceDir);
  }

  await Promise.all(tasks);
  // save to cache if exists
  if (cache && tasks) {
    await fs.mkdirs(path.dirname(cache));
    try {
      await new Promise((resolve, reject) => {
        zip.writeZip(cache, (err) => {
          if (err) {
            reject(err);
          } else {
            resolve({});
          }
        });
      });
    } catch (e) {
      if (e instanceof Error && (e as any)?.code === "ERR_OUT_OF_RANGE") {
        throw new ZipFileError(e);
      }
    }
  }
  return fs.createReadStream(cache);
}

export async function forEachFileAndDir(
  root: string,
  callback: (itemPath: string, stats: fs.Stats) => boolean | void,
  filter?: (itemPath: string) => boolean
): Promise<void> {
  await new Promise((resolve, reject) => {
    const stream: klaw.Walker = klaw(root, { filter: filter });
    stream
      .on("data", (item) => {
        if (callback(item.path, item.stats)) {
          stream.emit("close");
        }
      })
      .on("end", () => resolve({}))
      .on("error", (err) => reject(err))
      .on("close", () => resolve({}));
  });
}

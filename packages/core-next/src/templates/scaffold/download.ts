// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs";
import * as path from "path";
import AdmZip from "adm-zip";
import axios from "axios";
import type { ScaffoldContext, TemplateConfig } from "./types";

const DEFAULT_TRY_LIMITS = 1;
const DEFAULT_TIMEOUT_MS = 20000;

/**
 * Resolve the download URL for a template zip archive.
 * Returns undefined if version is "local" (use fallback only).
 */
export function resolveTemplateUrl(config: TemplateConfig, language: string): string | undefined {
  if (config.version === "local") {
    return undefined;
  }
  return `${config.downloadBaseUrl}/${config.tagPrefix}${config.version}/${language}${config.archiveExt}`;
}

/**
 * Download a zip archive from a URL with retry support.
 * Returns the Buffer on success, undefined on failure.
 */
export async function fetchZip(
  url: string,
  tryLimits: number = DEFAULT_TRY_LIMITS,
  timeoutMs: number = DEFAULT_TIMEOUT_MS
): Promise<Buffer | undefined> {
  for (let attempt = 0; attempt < tryLimits; attempt++) {
    try {
      const response = await axios.get(url, {
        responseType: "arraybuffer",
        timeout: timeoutMs,
      });
      return Buffer.from(response.data as ArrayBuffer);
    } catch {
      // Retry on failure
    }
  }
  return undefined;
}

/**
 * Extract a zip archive to a destination directory, applying file name and data transforms.
 * Returns a list of files written.
 */
export async function unzipWithTransform(
  zipBuffer: Buffer,
  ctx: ScaffoldContext
): Promise<string[]> {
  const zip = new AdmZip(zipBuffer);
  const entries = zip.getEntries();
  const written: string[] = [];

  for (const entry of entries) {
    if (entry.isDirectory) continue;

    const originalName = entry.entryName;

    // Zip Slip guard: reject entries containing path traversal sequences
    if (originalName.indexOf("..") !== -1) {
      throw new Error(
        `Zip Slip detected: entry "${originalName}" contains path traversal sequence`
      );
    }

    // Apply filter
    if (ctx.filterFn && !ctx.filterFn(originalName)) {
      continue;
    }

    const data = entry.getData();

    // Transform file name
    const outputName = ctx.fileNameReplaceFn
      ? ctx.fileNameReplaceFn(originalName, data)
      : originalName;

    // Transform file data
    const outputData = ctx.fileDataReplaceFn ? ctx.fileDataReplaceFn(originalName, data) : data;

    const outputPath = path.join(ctx.destination, outputName);

    // Zip Slip guard: ensure resolved path stays within the destination directory
    const resolvedOutput = path.resolve(outputPath);
    const resolvedDest = path.resolve(ctx.destination);
    if (!resolvedOutput.startsWith(resolvedDest + path.sep) && resolvedOutput !== resolvedDest) {
      throw new Error(
        `Zip Slip detected: entry "${originalName}" resolves outside destination directory`
      );
    }

    await fs.promises.mkdir(path.dirname(outputPath), { recursive: true });

    if (typeof outputData === "string") {
      await fs.promises.writeFile(outputPath, outputData, "utf-8");
    } else {
      await fs.promises.writeFile(outputPath, outputData);
    }
    written.push(outputPath);
  }

  return written;
}

/**
 * Load a local fallback zip from templates/fallback/{language}.zip.
 * Returns the zip Buffer or undefined if the file doesn't exist.
 */
export async function loadLocalFallback(
  fallbackDir: string,
  language: string
): Promise<Buffer | undefined> {
  const zipPath = path.join(fallbackDir, `${language}.zip`);
  try {
    return await fs.promises.readFile(zipPath);
  } catch {
    return undefined;
  }
}

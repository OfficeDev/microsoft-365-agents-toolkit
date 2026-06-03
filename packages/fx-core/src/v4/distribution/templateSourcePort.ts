// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError } from "@microsoft/teamsfx-api";
import axios, { AxiosResponse } from "axios";
import { Result, err, ok } from "neverthrow";
import * as fs from "fs-extra";
import * as path from "path";
import os from "os";
import { sendRequestWithRetry } from "../../common/requestUtils";
import {
  BundledFloor,
  CachedPackage,
  TagEntry,
  TemplateSource,
  TemplateSourcePort,
  computeDigest,
} from "./templateSource";

/**
 * Production wiring of the narrow {@link TemplateSourcePort} over real
 * environment / network / filesystem. The pure helpers (`parseTagList`,
 * `templateZipUrl`, `cacheDir`) are exported and unit-tested; the IO faces
 * are thin wrappers around them.
 *
 * Spec: docs/03-specs/operations/scaffolding/resolve-template-source.md
 * (decisions #6 digest = sha256 of zip bytes, #7 NDJSON tag list).
 */

const SOURCE = "Scaffold";

/** v4 channel prefix and naming (ADR-0006 channel isolation). */
// eslint-disable-next-line no-secrets/no-secrets
export const V4_TAG_PREFIX = "templates-v4@";
export const ZIP_EXT = ".zip";

export interface TemplateChannelConfig {
  /** NDJSON tag-list URL for the v4 channel (separate from the frozen v3 `tagListURL`). */
  templatesV4TagListURL: string;
  /** Base URL release assets are downloaded from. */
  templateDownloadBaseURL: string;
  /** Network retry budget. */
  tryLimits: number;
}

/**
 * Parse the model-A NDJSON tag list: one `{ version, digest }` object per
 * line. Blank lines and trailing `\r` are ignored; a malformed line is a hard
 * error (no silent skip) — spec decision #7.
 */
export function parseTagList(ndjson: string): TagEntry[] {
  const entries: TagEntry[] = [];
  const lines = ndjson.split("\n");
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].replace(/\r$/, "").trim();
    if (line.length === 0) {
      continue;
    }
    let parsed: unknown;
    try {
      parsed = JSON.parse(line);
    } catch {
      throw new SystemError({
        source: SOURCE,
        name: "TemplateTagListMalformed",
        message: `Malformed v4 tag-list entry on line ${i + 1}: not valid JSON.`,
      });
    }
    if (!isTagEntry(parsed)) {
      throw new SystemError({
        source: SOURCE,
        name: "TemplateTagListMalformed",
        message: `Malformed v4 tag-list entry on line ${i + 1}: missing "version" or "digest".`,
      });
    }
    entries.push({ version: parsed.version, digest: parsed.digest });
  }
  return entries;
}

function isTagEntry(value: unknown): value is TagEntry {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  const v = value as Record<string, unknown>;
  return typeof v.version === "string" && typeof v.digest === "string";
}

/** The download URL for a v4 template package version. */
export function templateZipUrl(baseURL: string, version: string): string {
  return `${baseURL}/${V4_TAG_PREFIX}${version}/templates${ZIP_EXT}`;
}

/** The on-disk cache directory for the v4 template packages. */
export function cacheDir(): string {
  return path.join(os.homedir(), ".fx", "templates-cache");
}

/** The cache file path for one version. */
export function cacheFile(version: string): string {
  return path.join(cacheDir(), `${V4_TAG_PREFIX}${version}${ZIP_EXT}`);
}

/**
 * Build the production port. `floor` is injected (it comes from the engine's
 * baked bundled package); this factory does not fabricate it.
 */
export function createTemplateSourcePort(
  config: TemplateChannelConfig,
  floor: BundledFloor
): TemplateSourcePort {
  const memo = new Map<string, CachedPackage>();

  // Warm the in-memory index from disk so `keys()` reflects prior downloads
  // for the offline `max(cache, floor)` fallback.
  let warmed = false;
  const warm = (): void => {
    if (warmed) {
      return;
    }
    warmed = true;
    let names: string[] = [];
    try {
      names = fs.readdirSync(cacheDir());
    } catch {
      return; // EAFP: no cache dir yet → empty index
    }
    for (const name of names) {
      if (!name.startsWith(V4_TAG_PREFIX) || !name.endsWith(ZIP_EXT)) {
        continue;
      }
      const version = name.slice(V4_TAG_PREFIX.length, name.length - ZIP_EXT.length);
      try {
        const bytes = fs.readFileSync(path.join(cacheDir(), name));
        memo.set(version, { digest: computeDigest(bytes), bytes });
      } catch {
        /* istanbul ignore next -- defensive: a listed cache file vanishing/locked between readdir and read is not reproducible */
        continue; // a vanished/locked file is simply not indexed
      }
    }
  };

  return {
    env: (name: string): string | undefined => process.env[name],

    tagList: async (): Promise<TagEntry[]> => {
      const res: AxiosResponse<string> = await sendRequestWithRetry(
        () => axios.get(config.templatesV4TagListURL, { responseType: "text" }),
        config.tryLimits
      );
      return parseTagList(res.data);
    },

    packages: async (version: string): Promise<Buffer> => {
      const res: AxiosResponse<ArrayBuffer> = await sendRequestWithRetry(
        () =>
          axios.get(templateZipUrl(config.templateDownloadBaseURL, version), {
            responseType: "arraybuffer",
          }),
        config.tryLimits
      );
      return Buffer.from(res.data);
    },

    cache: {
      get: (version: string): CachedPackage | undefined => {
        warm();
        return memo.get(version);
      },
      put: (version: string, digest: string, bytes: Buffer): void => {
        fs.ensureDirSync(cacheDir());
        fs.writeFileSync(cacheFile(version), bytes);
        memo.set(version, { digest, bytes });
      },
      keys: (): string[] => {
        warm();
        return [...memo.keys()];
      },
    },

    floor,
  };
}

/**
 * Load the bytes for an already-resolved {@link TemplateSource}.
 *
 * `resolveTemplateSource` decides *which* package a scaffold run uses but
 * returns only the descriptor; this is the companion IO step that produces the
 * bytes, the boundary the v3 chokepoint calls before `openTemplatePackage`.
 *
 * It never reaches the network: an `online`/`cache` source was already
 * downloaded, verified, and cached during resolution, so re-downloading here
 * would defeat the digest-keyed cache and reintroduce the v3 re-fetch churn. A
 * cache miss is therefore an invariant violation (hard error), not a silent
 * re-fetch. Every path re-checks the bytes against `source.digest` so
 * `computeDigest` stays the single integrity authority and a corrupt package is
 * never handed back.
 *
 * Spec: docs/03-specs/operations/scaffolding/resolve-template-source.md
 */
export function loadResolvedPackage(
  source: TemplateSource,
  port: TemplateSourcePort
): Result<Buffer, FxError> {
  let bytes: Buffer;
  if (source.origin === "bundled" || source.origin === "bundled-fallback") {
    try {
      bytes = fs.readFileSync(source.location);
    } catch {
      return err(
        new SystemError({
          source: SOURCE,
          name: "TemplatePackageUnreadable",
          message: `The resolved bundled template package at "${source.location}" is missing or unreadable.`,
        })
      );
    }
  } else {
    const cached = port.cache.get(source.version);
    if (cached === undefined) {
      return err(
        new SystemError({
          source: SOURCE,
          name: "TemplatePackageNotCached",
          message: `The resolved template package "${source.version}" is not present in the local cache.`,
        })
      );
    }
    bytes = cached.bytes;
  }

  const digest = computeDigest(bytes);
  if (digest !== source.digest) {
    return err(
      new SystemError({
        source: SOURCE,
        name: "TemplateDigestMismatch",
        message: `The resolved template package "${source.version}" failed its integrity check: expected ${source.digest}, got ${digest}.`,
      })
    );
  }
  return ok(bytes);
}

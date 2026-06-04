// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError, UserError } from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";
import * as crypto from "crypto";
import semver from "semver";

/**
 * The v4 template-distribution resolution operation.
 *
 * Spec: docs/03-specs/operations/scaffolding/resolve-template-source.md
 * Decision: docs/02-architecture/adr/ADR-0006-template-distribution-channel.md
 *
 * This module is part of the v4 world. It imports no v3 symbol; v3 code may
 * call `resolveTemplateSource`, but nothing here is tailored for v3.
 */

const SOURCE = "Scaffold";

/** Where the resolved template-package bytes come from. */
export type TemplateOrigin = "bundled" | "online" | "cache" | "bundled-fallback";

/** A single channel-published version paired with its expected content digest (model A). */
export interface TagEntry {
  version: string;
  digest: string;
}

/** A cached template package: its content digest and the bytes themselves. */
export interface CachedPackage {
  digest: string;
  bytes: Buffer;
}

/** The bundled floor baked into the engine binary. */
export interface BundledFloor {
  version: string;
  digest: string;
  location: string;
}

/**
 * The narrow port `resolveTemplateSource` depends on (interface-segregation).
 * The full `ScaffoldRuntime` composes this later; this operation never reaches
 * for faces it does not use.
 */
export interface TemplateSourcePort {
  /** Read an environment variable (e.g. `TEMPLATE_VERSION`). */
  env(name: string): string | undefined;
  /** The channel's published `{ version, digest }` entries. */
  tagList(): Promise<TagEntry[]>;
  /** Download a package's bytes for a version (verified against `expectedDigest` by the caller). */
  packages(version: string, expectedDigest: string): Promise<Buffer>;
  /** The local digest-keyed package cache. */
  cache: {
    get(version: string): CachedPackage | undefined;
    put(version: string, digest: string, bytes: Buffer): void;
    keys(): string[];
  };
  /** The bundled floor baked into the engine. */
  floor: BundledFloor;
}

/** The single resolved source a scaffold run will use. */
export interface TemplateSource {
  origin: TemplateOrigin;
  version: string;
  digest: string;
  location: string;
  /** Set when the resolved source diverged from the intended online source (observable, never silent). */
  warning?: string;
}

export interface ResolveTemplateSourceInput {
  /** SemVer range the build is permitted to resolve within. */
  range: string;
  /** `true` for test/offline/daily builds (bundled floor); `false` for shipped builds (release channel). */
  bundled: boolean;
  port: TemplateSourcePort;
}

/** sha256 content hash of a package's bytes, prefixed `sha256:`. */
export function computeDigest(bytes: Buffer): string {
  return "sha256:" + crypto.createHash("sha256").update(bytes).digest("hex");
}

/**
 * Resolve `(range, bundled, port)` to exactly one `TemplateSource` before any
 * template is read or rendered. Pure with respect to its inputs and the
 * current tag-list state (INV-6). Returns `Result` rather than throwing,
 * consistent with the toolkit-wide neverthrow rule.
 */
export async function resolveTemplateSource(
  input: ResolveTemplateSourceInput
): Promise<Result<TemplateSource, FxError>> {
  const { range, bundled, port } = input;

  const templateVersion = port.env("TEMPLATE_VERSION");
  if (templateVersion === "local") {
    return ok(floorSource(port.floor)); // AC-02 (override beats bundled=false)
  }
  if (templateVersion) {
    return pinnedOnline(templateVersion, port); // AC-03
  }
  if (bundled) {
    return ok(floorSource(port.floor)); // AC-01, AC-12 (never sniff package.json#version)
  }
  return resolveOnline(range, port);
}

function floorSource(floor: BundledFloor): TemplateSource {
  return {
    origin: "bundled",
    version: floor.version,
    digest: floor.digest,
    location: floor.location,
  };
}

async function pinnedOnline(
  version: string,
  port: TemplateSourcePort
): Promise<Result<TemplateSource, FxError>> {
  let tags: TagEntry[];
  try {
    tags = await port.tagList();
  } catch (e) {
    // A pin never falls back (AC-16 intent): surface the channel failure as a
    // Result instead of letting the rejection escape (neverthrow contract).
    return err(asTagListError(e));
  }
  const entry = tags.find((t) => t.version === version);
  if (!entry) {
    return err(
      new UserError({
        source: SOURCE,
        name: "TemplatePinnedVersionNotFound",
        message: `Pinned template version "${version}" is not published on the release channel.`,
      })
    );
  }
  return fetchVerify(entry, port, "online");
}

async function resolveOnline(
  range: string,
  port: TemplateSourcePort
): Promise<Result<TemplateSource, FxError>> {
  let tags: TagEntry[];
  try {
    tags = await port.tagList();
  } catch (e) {
    // A malformed channel document is a hard error (spec decision #7 / AC-17),
    // distinct from an unreachable channel: it must never be masked as an
    // offline fallback. Only a genuine reachability failure falls back.
    if (isMalformedTagList(e)) {
      return err(e);
    }
    return ok(offlineFallback(range, port)); // AC-07, AC-08 (unreachable only)
  }

  const picked = semver.maxSatisfying(
    tags.map((t) => t.version),
    range
  ); // AC-09 (stable excludes -beta), AC-10 (range names -beta)

  if (!picked) {
    // Channel reachable but no version satisfies range (AC-14 / AC-15).
    if (semver.satisfies(port.floor.version, range)) {
      return ok(floorFallback(port.floor)); // AC-14
    }
    return err(
      new UserError({
        source: SOURCE,
        name: "TemplateVersionMismatch",
        message: `No published template version satisfies range "${range}", and the bundled floor "${port.floor.version}" does not satisfy it either. The engine and template versions are incompatible.`,
      })
    ); // AC-15
  }

  if (picked === port.floor.version) {
    return ok(floorSource(port.floor)); // AC-05 (floor is highest satisfier; no download)
  }

  const entry = tags.find((t) => t.version === picked);
  /* istanbul ignore if -- defensive: `picked` is drawn from `tags`, so a miss
     means the tag list mutated mid-resolution; not reproducible in a unit. */
  if (!entry) {
    return err(
      new SystemError({
        source: SOURCE,
        name: "TemplateTagListInconsistent",
        message: `Resolved version "${picked}" is no longer present in the tag list.`,
      })
    );
  }
  return fetchVerify(entry, port, "online"); // AC-04, AC-06, AC-11
}

/** A malformed channel tag-list (decision #7) is a hard error, distinct from an unreachable channel. */
function isMalformedTagList(e: unknown): e is SystemError {
  return e instanceof SystemError && e.name === "TemplateTagListMalformed";
}

/** Preserve an existing FxError; wrap any other tag-list failure as a SystemError (neverthrow contract). */
function asTagListError(e: unknown): FxError {
  if (e instanceof UserError || e instanceof SystemError) {
    return e;
  }
  /* istanbul ignore next -- the port rejects with an Error; the String(e)
     fallback is defensive for non-Error rejections, not reproducible in a unit. */
  const message = e instanceof Error ? e.message : String(e);
  return new SystemError({
    source: SOURCE,
    name: "TemplateTagListUnavailable",
    message: `Failed to read the template release channel: ${message}.`,
  });
}

/** Preserve an existing FxError; wrap any other download failure as a SystemError (neverthrow contract). */
function asDownloadError(e: unknown, version: string): FxError {
  if (e instanceof UserError || e instanceof SystemError) {
    return e;
  }
  /* istanbul ignore next -- the port rejects with an Error; the String(e)
     fallback is defensive for non-Error rejections, not reproducible in a unit. */
  const message = e instanceof Error ? e.message : String(e);
  return new SystemError({
    source: SOURCE,
    name: "TemplateDownloadFailed",
    message: `Failed to download template "${version}" from the release channel: ${message}.`,
  });
}

async function fetchVerify(
  entry: TagEntry,
  port: TemplateSourcePort,
  origin: "online"
): Promise<Result<TemplateSource, FxError>> {
  const cached = port.cache.get(entry.version);
  if (cached && cached.digest === entry.digest) {
    return ok({
      origin: "cache",
      version: entry.version,
      digest: entry.digest,
      location: cacheLocation(entry.version),
    }); // AC-06 (zero download)
  }

  let bytes: Buffer;
  try {
    bytes = await port.packages(entry.version, entry.digest);
  } catch (e) {
    // The download (sendRequestWithRetry) can reject; surface it as a Result
    // rather than letting it escape resolveTemplateSource (neverthrow contract).
    return err(asDownloadError(e, entry.version));
  }
  const computed = computeDigest(bytes);
  if (computed !== entry.digest) {
    return err(
      new SystemError({
        source: SOURCE,
        name: "TemplateDigestMismatch",
        message: `Downloaded template "${entry.version}" failed integrity check: expected ${entry.digest}, got ${computed}.`,
      })
    ); // AC-11 (cache not written, corrupt bytes never returned)
  }

  port.cache.put(entry.version, entry.digest, bytes);
  return ok({
    origin,
    version: entry.version,
    digest: entry.digest,
    location: cacheLocation(entry.version),
  }); // AC-04
}

/** Reached when the channel is unreachable: max(highest cached satisfying range, floor). */
function offlineFallback(range: string, port: TemplateSourcePort): TemplateSource {
  const cachedSatisfying = port.cache.keys().filter((v) => semver.satisfies(v, range));
  const highestCached = semver.maxSatisfying(cachedSatisfying, range);
  const floorSatisfies = semver.satisfies(port.floor.version, range);

  // Cache wins only when strictly higher than the floor; a tie goes to the floor (decision #2).
  if (highestCached && (!floorSatisfies || semver.gt(highestCached, port.floor.version))) {
    const cached = port.cache.get(highestCached);
    /* istanbul ignore else -- highestCached is drawn from cache.keys(); a miss
       means the cache mutated mid-resolution, so we degrade to the floor. */
    if (cached) {
      return {
        origin: "cache",
        version: highestCached,
        digest: cached.digest,
        location: cacheLocation(highestCached),
        warning: offlineWarning(range),
      }; // AC-07
    }
  }

  return { ...floorFallback(port.floor), warning: offlineWarning(range) }; // AC-08
}

function floorFallback(floor: BundledFloor): TemplateSource {
  return {
    origin: "bundled-fallback",
    version: floor.version,
    digest: floor.digest,
    location: floor.location,
    warning: `Falling back to the bundled template floor "${floor.version}".`,
  };
}

function offlineWarning(range: string): string {
  return `The template release channel was unreachable; resolved offline within range "${range}".`;
}

function cacheLocation(version: string): string {
  // eslint-disable-next-line no-secrets/no-secrets
  return `templates-v4@${version}`;
}

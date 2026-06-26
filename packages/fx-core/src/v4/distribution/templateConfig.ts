// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as semver from "semver";

/** Build-time v4 template distribution config. See resolve-template-source spec. */

export interface V4TemplateConfigInput {
  /** Freshly minted templates version. */
  version: string;
  /** CD `goproduct` flag: `true` when this build ships to users. */
  goproduct: boolean;
  /** The current `range` in `templates-config.json`, kept when still satisfied. */
  previousRange: string;
}

export interface V4TemplateConfig {
  /** `~major.minor` SemVer range the build may resolve within. */
  range: string;
  /** `true` for test/offline builds (bundled floor); `false` for shipped builds. */
  bundled: boolean;
  /** The clean, suffix-free version published to the v4 channel (floor pin). */
  localVersion: string;
}

/** Non-shipping builds resolve from the bundled floor; shipping builds resolve online. */
export function computeBundled(goproduct: boolean): boolean {
  return !goproduct;
}

/** Compute the clean, suffix-free version published to the v4 channel. */
export function computeV4PublishVersion(version: string): string {
  const parsed = semver.parse(version);
  if (parsed === null) {
    throw new Error(`Cannot compute v4 publish version: "${version}" is not valid SemVer.`);
  }
  if (parsed.minor % 2 === 1) {
    const dateStamp = parsed.prerelease.find(
      (segment): segment is number => typeof segment === "number" && segment >= 1_000_000_000
    );
    if (dateStamp !== undefined) {
      return `${parsed.major}.${parsed.minor}.${dateStamp}`;
    }
  }
  return `${parsed.major}.${parsed.minor}.${parsed.patch}`;
}

/** Compute the SemVer range a build may resolve within. */
export function computeRange(version: string, previousRange: string): string {
  const parsed = semver.parse(version);
  if (parsed === null) {
    throw new Error(`Cannot compute v4 template range: "${version}" is not valid SemVer.`);
  }
  if (semver.intersects(previousRange, `${parsed.major}.${parsed.minor}.0`)) {
    return previousRange;
  }
  return `~${parsed.major}.${parsed.minor}`;
}

/** Compute the full v4 distribution config block for `templates-config.json`. */
export function computeV4TemplateConfig(input: V4TemplateConfigInput): V4TemplateConfig {
  return {
    range: computeRange(input.version, input.previousRange),
    bundled: computeBundled(input.goproduct),
    localVersion: computeV4PublishVersion(input.version),
  };
}

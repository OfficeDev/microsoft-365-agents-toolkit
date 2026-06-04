// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as semver from "semver";

/**
 * Build-time computation of the v4 template distribution config from the freshly
 * minted templates version and the CD `goproduct` flag.
 *
 * This is the v4-isolated counterpart of the v3 `fxcore-sync-up-version.js`
 * (`syncTemplateVersion` + `updateUseLocalFlag`). It is a pure function so it
 * can be unit-tested 1:1; the CD step is a thin node wrapper that reads the
 * version + flag, calls this, and writes the result into `templates-config.json`.
 *
 * Outputs the two build-time concepts `resolveTemplateSource` consumes:
 *   - `range`   — the SemVer range the build may resolve within (`~major.minor`)
 *   - `bundled` — `true` for test/offline builds (bundled floor), `false` for
 *                 shipped builds (release channel)
 * plus `localVersion` (the exact minted version, recorded for observability).
 *
 * Spec: docs/03-specs/operations/scaffolding/resolve-template-source.md
 */

export interface V4TemplateConfigInput {
  /** The freshly minted templates version (e.g. `6.11.0` or `6.11.0-rc.0`). */
  version: string;
  /**
   * The CD `goproduct` flag (env `PRODUCTION`): `true` when this build ships to
   * users (marketplace / stable), `false` for any internal test build.
   */
  goproduct: boolean;
  /** The current `range` in `templates-config.json`, kept when still satisfied. */
  previousRange: string;
}

export interface V4TemplateConfig {
  /** `~major.minor` SemVer range the build may resolve within. */
  range: string;
  /** `true` for test/offline builds (bundled floor); `false` for shipped builds. */
  bundled: boolean;
  /** The exact minted version (observability / floor pin). */
  localVersion: string;
}

/**
 * `bundled` is the negation of `goproduct`: a build that is NOT shipping to
 * users resolves from the bundled floor (offline-by-default); a shipping build
 * resolves from the online release channel. Independent of `preid` / lane.
 */
export function computeBundled(goproduct: boolean): boolean {
  return !goproduct;
}

/**
 * The range a build may resolve within. templates carries no odd/even-minor
 * split (unlike the VS Code extension), so a prerelease and the stable it
 * graduates into share one `~major.minor`. The range is only widened when the
 * minted version no longer intersects the previous range; otherwise it is kept
 * stable for reproducibility.
 */
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
    localVersion: input.version,
  };
}

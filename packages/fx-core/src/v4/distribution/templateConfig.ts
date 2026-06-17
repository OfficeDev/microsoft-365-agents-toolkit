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
 * plus `localVersion` (the clean, suffix-free version published to / pinned on
 * the v4 channel — see `computeV4PublishVersion`).
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
  /** The clean, suffix-free version published to the v4 channel (floor pin). */
  localVersion: string;
}

/**
 * `bundled` decides whether a build resolves from the bundled floor
 * (offline-by-default) or the online release channel: it is simply the negation
 * of `goproduct`. A build that is NOT shipping to users (internal / daily test
 * builds) resolves from the bundled floor; a shipping build resolves online.
 *
 * A prerelease-SUFFIXED minted version does NOT force bundling: a shipping
 * prerelease (odd-minor preview, `goproduct=true`) is published online under a
 * clean, suffix-free version computed by `computeV4PublishVersion` so a
 * `~major.minor` range can resolve it.
 */
export function computeBundled(goproduct: boolean): boolean {
  return !goproduct;
}

/**
 * The clean, suffix-free version published to the v4 channel for a minted
 * version, derived purely from the minted version's odd/even minor:
 *
 *   - ODD minor  (prerelease line, e.g. `6.11.x`): the build date stamped into
 *     the patch (`6.11.<YYYYMMDDHH>`). The date is read from the `-beta.<date>`
 *     preid the preview lane mints, so every preview build gets a unique,
 *     suffix-free version that satisfies `~major.minor`. Falls back to the raw
 *     patch when no date stamp is present.
 *   - EVEN minor (stable line, e.g. `6.10.x`): `major.minor.patch` as-is; the
 *     stable lane already mints a clean version.
 *
 * This is the v4-channel counterpart of the clean `templates@<major>.<minor>.<date>`
 * pattern the v3 channel uses for prereleases. Refusing to publish an
 * even-minor SUFFIXED version (a preview lane minted on a stable branch) is
 * enforced at publish time, not here, so non-shipping test builds keep working.
 */
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

/**
 * The range a build may resolve within: `~major.minor` of the minted version.
 * v4 templates use an odd/even-minor split — odd-minor prereleases (`6.11.x`)
 * and even-minor stables (`6.10.x`) live on different minors, so their ranges
 * are naturally isolated. The range is only widened when the minted version no
 * longer intersects the previous range; otherwise it is kept stable for
 * reproducibility.
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
    localVersion: computeV4PublishVersion(input.version),
  };
}

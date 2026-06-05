// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { featureFlagManager, FeatureFlags } from "../../common/featureFlags";

const packageJson = require("../../../package.json");

/**
 * Determines whether to use local templates based on environment variables and package version.
 * Returns true if:
 * - TEMPLATE_VERSION env variable is set to "local", OR
 * - Package version contains "alpha" (daily build version)
 */
export function useLocalTemplate(): boolean {
  const templateVersionEnv = process.env["TEMPLATE_VERSION"];
  if (templateVersionEnv === "local") {
    return true;
  }
  const version: string = packageJson.version;
  if (version.includes("alpha")) {
    // daily build version
    return true;
  }

  return false;
}

/**
 * Transitional: the v4 channel only publishes online metadata for stable
 * (goproduct) releases under `templates-v4@<ver>`. Prerelease/test builds have
 * no v4 release, so when the v4 flag is on for a prerelease build the metadata
 * and UI readers must read the bundled copy and ignore any (possibly stale v3)
 * `~/.fx` cache left by a previous v3 run. Stable v4 builds download v4 metadata
 * into `~/.fx` and read it from there as usual.
 *
 * Remove once selector.json drives metadata distribution.
 */
export function useBundledMetadataForV4(): boolean {
  if (!featureFlagManager.getBooleanValue(FeatureFlags.V4Enabled)) {
    return false;
  }
  const version: string = packageJson.version;
  return version.includes("alpha") || version.includes("beta") || version.includes("rc");
}

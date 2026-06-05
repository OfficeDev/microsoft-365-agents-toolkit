// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ConfigFolderName } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import os from "os";
import path from "path";
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
 * Transitional: in the v4 channel the metadata/UI readers must read the bundled
 * copy and ignore any (possibly stale v3) `~/.fx` cache UNLESS the v4 online
 * fetch populated its cache. `fetchOnlineTemplateMetadata` writes
 * `~/.fx/template-version-v4.txt` only after a successful v4 download, so its
 * presence is the single signal that downloaded v4 metadata is available:
 *   - present → read the downloaded v4 cache;
 *   - absent  → bundled build / channel unreachable / not yet published, so
 *               read the bundled copy (never a stale v3 cache).
 * This mirrors the `resolveTemplateSource` decision in
 * `fetchOnlineTemplateMetadata`.
 *
 * Remove once selector.json drives metadata distribution.
 */
export function useBundledMetadataForV4(): boolean {
  if (!featureFlagManager.getBooleanValue(FeatureFlags.V4Enabled)) {
    return false;
  }
  const v4VersionFile = path.join(
    os.homedir(),
    `.${String(ConfigFolderName)}`,
    "template-version-v4.txt"
  );
  return !fs.pathExistsSync(v4VersionFile);
}

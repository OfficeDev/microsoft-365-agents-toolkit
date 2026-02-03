// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

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

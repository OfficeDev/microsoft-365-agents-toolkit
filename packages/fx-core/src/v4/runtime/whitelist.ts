// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ExpressionRuntimePort, WhitelistFn } from "../expression/evaluateExpression";
import { readBooleanFeatureFlag } from "../../common/featureFlags";

import { safeProjectNameLowerCase, pathDelimiter, contains } from "./functions/generic";
export { safeProjectNameLowerCase, pathDelimiter, contains } from "./functions/generic";
import { mcpNamespace, mcpAuthRef } from "./functions/mcp";
export { deriveMcpServerName, mcpNamespace, mcpAuthRef } from "./functions/mcp";
import {
  officeAddinManifestScope,
  officeAddinManifestScopes,
  officeAddinLaunchConfigurations,
  officeAddinLaunchCompounds,
  officeAddinDebugScripts,
  officeAddinDebugApp,
} from "./functions/officeAddin";
export {
  officeAddinManifestScope,
  officeAddinManifestScopes,
  officeAddinLaunchConfigurations,
  officeAddinLaunchCompounds,
  officeAddinDebugScripts,
  officeAddinDebugApp,
} from "./functions/officeAddin";

/** The closed function whitelist; an author cannot extend it (ADR-0016 decision 3). */
const WHITELIST = new Map<string, WhitelistFn>([
  ["mcpNamespace", mcpNamespace],
  ["mcpAuthRef", mcpAuthRef],
  ["safeProjectNameLowerCase", safeProjectNameLowerCase],
  ["pathDelimiter", pathDelimiter],
  ["contains", contains],
  ["officeAddinManifestScope", officeAddinManifestScope],
  ["officeAddinManifestScopes", officeAddinManifestScopes],
  ["officeAddinLaunchConfigurations", officeAddinLaunchConfigurations],
  ["officeAddinLaunchCompounds", officeAddinLaunchCompounds],
  ["officeAddinDebugScripts", officeAddinDebugScripts],
  ["officeAddinDebugApp", officeAddinDebugApp],
]);

/** Default feature-flag reader: an env-backed truthy check (`"true"` / `"1"`). */
function envFlagReader(name: string): boolean {
  return readBooleanFeatureFlag(name);
}

/** Build the real pure expression runtime port. */
export function createExpressionPort(
  flagReader: (name: string) => boolean = envFlagReader
): ExpressionRuntimePort {
  return {
    functions: (name: string): WhitelistFn | undefined => WHITELIST.get(name),
    flags: (name: string): boolean => flagReader(name),
  };
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { DriverStep, LifecycleAnalysis } from "./types";
import { resolveConfig } from "./resolver";

/**
 * Driver IDs that require an M365 (Teams / Entra ID) authentication scope.
 */
const M365_DRIVERS = new Set([
  "teamsApp/create",
  "teamsApp/configure",
  "teamsApp/update",
  "teamsApp/zipAppPackage",
  "teamsApp/validateManifest",
  "teamsApp/validateAppPackage",
  "teamsApp/publishAppPackage",
  "teamsApp/extendToM365",
  "aadApp/create",
  "aadApp/update",
  "botAadApp/create",
  "botFramework/create",
  "oauth/register",
  "apiKey/register",
]);

/**
 * Driver IDs that require an Azure subscription / credentials.
 */
const AZURE_DRIVERS = new Set([
  "arm/deploy",
  "azureAppService/zipDeploy",
  "azureFunctions/zipDeploy",
  "azureStorage/deploy",
  "azureStorage/config",
]);

/**
 * Regex matching ${{VAR_NAME}} placeholders in string values.
 */
const PLACEHOLDER_RE = /\$\{\{([A-Za-z_][A-Za-z0-9_]*)\}\}/g;

/**
 * Analyze a set of lifecycle steps to determine what prerequisites are needed.
 *
 * Scans driver IDs to check for M365 or Azure auth requirements and
 * collects unresolved placeholder variables from step configs.
 *
 * @param steps   The driver steps to analyze.
 * @param envMap  Optional current environment (to detect unresolved placeholders).
 * @returns Analysis result with auth requirements and unresolved vars.
 */
export function analyzeSteps(
  steps: DriverStep[],
  envMap?: ReadonlyMap<string, string> | Record<string, string>
): LifecycleAnalysis {
  let needsM365 = false;
  let needsAzure = false;
  const driverIds: string[] = [];
  const unresolvedVars: string[] = [];

  for (const step of steps) {
    driverIds.push(step.uses);

    if (M365_DRIVERS.has(step.uses)) {
      needsM365 = true;
    }
    if (AZURE_DRIVERS.has(step.uses)) {
      needsAzure = true;
    }

    // Collect unresolved placeholders from step config
    if (envMap) {
      const { unresolved } = resolveConfig(step.with, envMap);
      for (const u of unresolved) {
        if (!unresolvedVars.includes(u.name)) {
          unresolvedVars.push(u.name);
        }
      }
    } else {
      // Without envMap, collect all placeholders as potentially unresolved
      collectPlaceholders(step.with, unresolvedVars);
    }
  }

  return { needsM365, needsAzure, driverIds, unresolvedVars };
}

/**
 * Collect all ${{VAR}} placeholder names from an object tree.
 */
function collectPlaceholders(obj: unknown, found: string[]): void {
  if (typeof obj === "string") {
    let match: RegExpExecArray | null;
    while ((match = PLACEHOLDER_RE.exec(obj)) !== null) {
      if (!found.includes(match[1])) {
        found.push(match[1]);
      }
    }
  } else if (Array.isArray(obj)) {
    for (const item of obj) {
      collectPlaceholders(item, found);
    }
  } else if (obj !== null && typeof obj === "object") {
    for (const value of Object.values(obj)) {
      collectPlaceholders(value, found);
    }
  }
}

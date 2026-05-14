// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { driverRegistry } from "../registry";
import { createOrUpdateEnvironmentFileDriver } from "./file/createOrUpdateEnvironmentFile";
import { createOrUpdateJsonFileDriver } from "./file/createOrUpdateJsonFile";
import { scriptDriver } from "./script/run";
import { zipAppPackageDriver } from "./teamsApp/zipAppPackage";
import { validateManifestDriver } from "./teamsApp/validateManifest";
import { validateAppPackageDriver } from "./teamsApp/validateAppPackage";
import { createTeamsAppDriver } from "./teamsApp/create";
import { configureTeamsAppDriver } from "./teamsApp/configure";
import { publishAppPackageDriver } from "./teamsApp/publishAppPackage";
import { createAadAppDriver } from "./aadApp/create";
import { updateAadAppDriver } from "./aadApp/update";
import { createBotAadAppDriver } from "./botAadApp/create";
import { createBotFrameworkDriver } from "./botFramework/create";
import { armDeployDriver } from "./arm/deploy";
import { azureAppServiceZipDeployDriver } from "./azureAppService/zipDeploy";
import { azureFunctionsZipDeployDriver } from "./azureFunctions/zipDeploy";
import { oauthRegisterDriver } from "./oauth/register";
import { apiKeyRegisterDriver } from "./apiKey/register";
import { updateTeamsAppDriver } from "./teamsApp/update";
import { extendToM365Driver } from "./teamsApp/extendToM365";
import { runNpmCommandDriver } from "./cli/runNpmCommand";
import { runDotnetCommandDriver } from "./cli/runDotnetCommand";

/** All built-in driver descriptors */
export const builtinDrivers = [
  createOrUpdateEnvironmentFileDriver,
  createOrUpdateJsonFileDriver,
  scriptDriver,
  zipAppPackageDriver,
  validateManifestDriver,
  validateAppPackageDriver,
  createTeamsAppDriver,
  configureTeamsAppDriver,
  publishAppPackageDriver,
  createAadAppDriver,
  updateAadAppDriver,
  createBotAadAppDriver,
  createBotFrameworkDriver,
  armDeployDriver,
  azureAppServiceZipDeployDriver,
  azureFunctionsZipDeployDriver,
  oauthRegisterDriver,
  apiKeyRegisterDriver,
  updateTeamsAppDriver,
  extendToM365Driver,
  runNpmCommandDriver,
  runDotnetCommandDriver,
] as const;

/**
 * Register all built-in drivers with the global DriverRegistry.
 * Call once at application startup before executing any lifecycle.
 */
export function registerBuiltinDrivers(): void {
  for (const driver of builtinDrivers) {
    if (!driverRegistry.has(driver.id)) {
      driverRegistry.register(driver);
    }
  }
}

// Re-export individual drivers for direct import
export { createOrUpdateEnvironmentFileDriver } from "./file/createOrUpdateEnvironmentFile";
export { createOrUpdateJsonFileDriver } from "./file/createOrUpdateJsonFile";
export { scriptDriver } from "./script/run";
export { zipAppPackageDriver } from "./teamsApp/zipAppPackage";
export { validateManifestDriver } from "./teamsApp/validateManifest";
export { validateAppPackageDriver } from "./teamsApp/validateAppPackage";
export { createTeamsAppDriver } from "./teamsApp/create";
export { configureTeamsAppDriver } from "./teamsApp/configure";
export { publishAppPackageDriver } from "./teamsApp/publishAppPackage";
export { createAadAppDriver } from "./aadApp/create";
export { updateAadAppDriver } from "./aadApp/update";
export { createBotAadAppDriver } from "./botAadApp/create";
export { createBotFrameworkDriver } from "./botFramework/create";
export { armDeployDriver } from "./arm/deploy";
export { azureAppServiceZipDeployDriver } from "./azureAppService/zipDeploy";
export { azureFunctionsZipDeployDriver } from "./azureFunctions/zipDeploy";
export { oauthRegisterDriver } from "./oauth/register";
export { apiKeyRegisterDriver } from "./apiKey/register";
export { updateTeamsAppDriver } from "./teamsApp/update";
export { extendToM365Driver } from "./teamsApp/extendToM365";
export { runNpmCommandDriver } from "./cli/runNpmCommand";

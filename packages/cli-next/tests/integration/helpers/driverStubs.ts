/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 *
 * Provides sinon stubs for external-call drivers so that lifecycle
 * integration tests can run without real Azure/M365/npm calls.
 *
 * Design:
 * - File/env drivers (`file/*`) are left real — they write to the temp dir
 * - Script driver is stubbed (no real shell exec in CI)
 * - All service-call drivers return `ok({ outputs })` with mock data
 * - Each stubbed driver returns outputs matching what templates expect
 *   in their `writeToEnvironmentFile` sections
 */

import * as sinon from "sinon";
import { ok } from "neverthrow";
import { driverRegistry } from "@microsoft/teamsfx-core-next";
import type { DriverOutput } from "@microsoft/teamsfx-core-next";

/**
 * Default mock outputs per driver ID.
 * These match the `writeToEnvironmentFile` keys used in template YAML files.
 */
const MOCK_OUTPUTS: Record<string, Record<string, string>> = {
  "teamsApp/create": {
    teamsAppId: "mock-teams-app-id-00000",
  },
  "teamsApp/configure": {
    teamsAppId: "mock-teams-app-id-00000",
  },
  "teamsApp/update": {
    teamsAppId: "mock-teams-app-id-00000",
  },
  "teamsApp/zipAppPackage": {},
  "teamsApp/validateManifest": {},
  "teamsApp/validateAppPackage": {},
  "teamsApp/publishAppPackage": {
    publishedAppId: "mock-published-app-id-11111",
  },
  "teamsApp/extendToM365": {
    titleId: "mock-title-id-22222",
    appId: "mock-m365-app-id-33333",
    shareLink: "https://teams.microsoft.com/l/app/mock-teams-app-id-00000",
  },
  "arm/deploy": {
    BOT_AZURE_APP_SERVICE_RESOURCE_ID:
      "/subscriptions/00000000/resourceGroups/rg-test/providers/Microsoft.Web/sites/test-bot",
    AZURE_FUNCTION_RESOURCE_ID:
      "/subscriptions/00000000/resourceGroups/rg-test/providers/Microsoft.Web/sites/test-func",
    SECRET_BOT_PASSWORD: "mock-bot-password",
    BOT_DOMAIN: "test-bot.azurewebsites.net",
    BOT_ENDPOINT: "https://test-bot.azurewebsites.net",
  },
  "azureAppService/zipDeploy": {},
  "azureFunctions/zipDeploy": {},
  "aadApp/create": {
    clientId: "mock-aad-client-id-44444",
    clientSecret: "mock-aad-client-secret",
    objectId: "mock-aad-object-id-55555",
    tenantId: "mock-tenant-id-66666",
    authority: "https://login.microsoftonline.com/mock-tenant-id-66666",
    authorityHost: "https://login.microsoftonline.com",
  },
  "aadApp/update": {},
  "botAadApp/create": {
    botId: "mock-bot-id-77777",
    botPassword: "mock-bot-password",
  },
  "botFramework/create": {},
  "oauth/register": {
    registrationId: "mock-oauth-reg-id-88888",
    configurationId: "mock-oauth-config-id-88888",
    applicationIdUri: "api://mock-oauth-app-id-uri",
  },
  "apiKey/register": {
    registrationId: "mock-apikey-reg-id-99999",
  },
  "cli/runNpmCommand": {},
  script: {},
};

/**
 * IDs of drivers whose `executeFn` should be stubbed.
 * File drivers are intentionally excluded — they run real against the temp dir.
 */
const STUB_DRIVER_IDS = new Set(Object.keys(MOCK_OUTPUTS));

/**
 * Stub all external-call drivers in the global `driverRegistry`.
 * Returns a sinon sandbox that must be `restore()`d in `afterEach`.
 *
 * Each stubbed driver's `executeFn` becomes a sinon stub that returns
 * `ok({ outputs: MOCK_OUTPUTS[driverId] })`.
 *
 * @param sandbox  The sinon sandbox to use for stubs.
 * @param overrides  Optional per-driver output overrides.
 */
export function stubExternalDrivers(
  sandbox: sinon.SinonSandbox,
  overrides?: Record<string, Record<string, string>>
): Map<string, sinon.SinonStub> {
  const stubs = new Map<string, sinon.SinonStub>();

  for (const driverId of STUB_DRIVER_IDS) {
    const descriptor = driverRegistry.get(driverId);
    if (!descriptor) continue; // driver not registered — skip

    const outputs = { ...(MOCK_OUTPUTS[driverId] ?? {}), ...(overrides?.[driverId] ?? {}) };
    const mockResult: DriverOutput = { outputs };

    const stub = sandbox.stub(descriptor, "executeFn").resolves(ok(mockResult));
    stubs.set(driverId, stub);
  }

  return stubs;
}

/**
 * Get the set of driver IDs that are stubbed (not run for real).
 * Useful for test assertions about which drivers executed.
 */
export function getStubbedDriverIds(): Set<string> {
  return new Set(STUB_DRIVER_IDS);
}

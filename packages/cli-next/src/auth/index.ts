// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { TokenProvider } from "@microsoft/teamsfx-core-next";
import { checkAzureSPFile } from "./cacheAccess";
import M365Login from "./m365Login";
import AzureAccountManager from "./azureLogin";
import AzureLoginCI from "./azureLoginCI";

export { M365Login, AzureAccountManager, AzureLoginCI };

/**
 * Create a real TokenProvider for CLI usage.
 *
 * Selects the Azure provider based on:
 * - Service principal file exists → AzureLoginCI (headless / CI)
 * - Otherwise → AzureAccountManager (interactive MSAL)
 */
export function createTokenProvider(): TokenProvider {
  const azureAccountProvider = checkAzureSPFile()
    ? AzureLoginCI.getInstance()
    : AzureAccountManager.getInstance();

  return {
    m365TokenProvider: M365Login.getInstance(),
    azureAccountProvider,
  };
}

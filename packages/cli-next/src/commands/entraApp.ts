// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Command } from "commander";
import { wrapHandler } from "../handler";

/**
 * Entra ID app commands: entra-app update
 */
export function createEntraAppCommands(program: Command): void {
  const entraApp = program
    .command("entra-app")
    .description("Manage Entra ID (Azure AD) app registrations");

  // atk entra-app update
  entraApp
    .command("update")
    .description("Update Entra ID app registration based on the manifest")
    .option("-e, --env <environment>", "Target environment", "dev")
    .option("--manifest-path <path>", "Path to the Entra app manifest file")
    .action(
      wrapHandler("entra-app update", async (_opts, _cmd) => {
        console.log("Updating Entra app...");
      })
    );
}

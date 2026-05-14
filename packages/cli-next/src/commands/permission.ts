// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Command } from "commander";
import { wrapHandler } from "../handler";

/**
 * Permission commands: permission grant, permission status
 */
export function createPermissionCommands(program: Command): void {
  const permission = program
    .command("permission")
    .description("Manage collaboration permissions for the project");

  // atk permission grant
  permission
    .command("grant")
    .description("Grant permission to a collaborator")
    .option("-e, --env <environment>", "Target environment", "dev")
    .option("--email <email>", "Email of the collaborator to grant access to")
    .action(
      wrapHandler("permission grant", async (_opts, _cmd) => {
        console.log("Granting permission...");
      })
    );

  // atk permission status
  permission
    .command("status")
    .description("Show the permission status of the project")
    .option("-e, --env <environment>", "Target environment", "dev")
    .action(
      wrapHandler("permission status", async (_opts, _cmd) => {
        console.log("Showing permission status...");
      })
    );
}

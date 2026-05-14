// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Command } from "commander";
import { wrapHandler, wrapHandlerWithContext } from "../handler";
import { m365SideloadAction } from "../actions/m365Sideload";

/**
 * M365 commands: m365-sideload, m365-unacquire, m365-launch-info
 */
export function createM365Commands(program: Command): void {
  // atk m365-sideload
  program
    .command("m365-sideload")
    .description("Sideload the app to Microsoft 365")
    .requiredOption("--file-path <path>", "Path to the app package file")
    .option("--scope <scope>", "Sideload scope: Personal, Shared, Tenant", "Personal")
    .action(
      wrapHandlerWithContext("m365-sideload", async (ctx, opts) => {
        const result = await m365SideloadAction(ctx, {
          filePath: opts.filePath as string,
          scope: opts.scope as string | undefined,
        });
        console.log(`Sideloaded to M365 — Title ID: ${result.titleId}, App ID: ${result.appId}`);
        if (result.shareLink) {
          console.log(`Share link: ${result.shareLink}`);
        }
      })
    );

  // atk m365-unacquire
  program
    .command("m365-unacquire")
    .description("Remove the app from Microsoft 365")
    .option("--manifest-id <id>", "App manifest ID to remove")
    .option("--title-id <id>", "Title ID of the app to remove")
    .action(
      wrapHandler("m365-unacquire", async (_opts, _cmd) => {
        console.log("Removing app from M365...");
      })
    );

  // atk m365-launch-info
  program
    .command("m365-launch-info")
    .description("Get launch information for the M365 app")
    .option("--manifest-id <id>", "App manifest ID")
    .option("--title-id <id>", "Title ID")
    .action(
      wrapHandler("m365-launch-info", async (_opts, _cmd) => {
        console.log("Getting launch info...");
      })
    );
}

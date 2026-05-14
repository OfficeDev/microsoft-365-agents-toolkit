// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Command } from "commander";
import { wrapHandler, wrapHandlerWithContext } from "../handler";
import { setSensitivityLabelAction } from "../actions/setSensitivityLabel";

/**
 * Miscellaneous commands: validate, set, share, init
 */
export function createMiscCommands(program: Command): void {
  // atk validate
  program
    .command("validate")
    .description("Validate the project configuration and manifest")
    .option("-e, --env <environment>", "Target environment", "dev")
    .option("--manifest-path <path>", "Path to the manifest file")
    .option("--app-package-file-path <path>", "Path to the app package zip file")
    .action(
      wrapHandler("validate", async (_opts, _cmd) => {
        console.log("Validating project...");
      })
    );

  // atk set
  const set = program.command("set").description("Set project configurations");

  // atk set sensitivityLabel
  set
    .command("sensitivityLabel")
    .description("Set the sensitivity label for the project")
    .requiredOption("--label-id <id>", "Sensitivity label ID to set")
    .option("--agent-manifest-path <path>", "Path to the DA manifest (auto-detected if omitted)")
    .action(
      wrapHandlerWithContext("set sensitivityLabel", async (ctx, opts) => {
        await setSensitivityLabelAction(ctx, {
          projectPath: ctx.projectPath ?? process.cwd(),
          agentManifestPath: opts.agentManifestPath as string | undefined,
          labelId: opts.labelId as string,
        });
        console.log("Sensitivity label set successfully.");
      })
    );

  // atk share
  const share = program
    .command("share")
    .description("Share your app with others")
    .option("-e, --env <environment>", "Target environment", "dev")
    .action(
      wrapHandler("share", async (_opts, _cmd) => {
        console.log("Sharing app...");
      })
    );

  // atk share remove
  share
    .command("remove")
    .description("Remove a shared user from the app")
    .option("-e, --env <environment>", "Target environment", "dev")
    .action(
      wrapHandler("share remove", async (_opts, _cmd) => {
        console.log("Removing shared user...");
      })
    );

  // atk init
  program
    .command("init")
    .description("Initialize an existing project for use with the toolkit")
    .option("-f, --folder <folder>", "Project folder to initialize", ".")
    .action(
      wrapHandler("init", async (_opts, _cmd) => {
        console.log("Initializing project...");
      })
    );
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Command } from "commander";
import { wrapHandler, wrapHandlerWithContext } from "../handler";
import { validateAction, packageAction } from "../actions/teamsapp";
import { colorize, TextType } from "../output";

/**
 * Teams app management commands: validate, package, publish, update, doctor
 */
export function createTeamsappCommands(program: Command): void {
  const teamsapp = program.command("teamsapp").description("Manage Teams app lifecycle");

  // atk teamsapp validate
  teamsapp
    .command("validate")
    .description("Validate the Teams app manifest and configuration")
    .option("--manifest-path <path>", "Path to the app manifest file")
    .option("--app-package-file-path <path>", "Path to the app package zip file")
    .action(
      wrapHandlerWithContext("teamsapp validate", async (ctx, opts) => {
        const result = await validateAction(ctx, {
          manifestPath: opts.manifestPath as string | undefined,
          packagePath: opts.appPackageFilePath as string | undefined,
        });
        if (result.valid) {
          console.log(colorize("Validation passed.", TextType.Success));
        } else {
          for (const error of result.errors) {
            console.error(colorize(`  Error: ${error}`, TextType.Error));
          }
          if (result.warnings) {
            for (const warning of result.warnings) {
              console.warn(colorize(`  Warning: ${warning}`, TextType.Warning));
            }
          }
          process.exitCode = 1;
        }
      })
    );

  // atk teamsapp package
  teamsapp
    .command("package")
    .description("Package the Teams app into a zip file")
    .option("--manifest-path <path>", "Path to the app manifest file")
    .option(
      "--output-zip-path <path>",
      "Output path for the zip file",
      "./appPackage/build/appPackage.zip"
    )
    .option("--project-folder <folder>", "Path to the project folder")
    .action(
      wrapHandlerWithContext("teamsapp package", async (ctx, opts) => {
        const result = await packageAction(ctx, {
          projectPath: ctx.projectPath ?? process.cwd(),
          manifestPath: opts.manifestPath as string | undefined,
          outputPath: opts.outputZipPath as string,
        });
        console.log(colorize(`App package created: ${result.zipPath}`, TextType.Success));
      })
    );

  // atk teamsapp publish (stub — delegated to lifecycle)
  teamsapp
    .command("publish")
    .description("Publish the Teams app to the organization's app catalog")
    .option("-e, --env <environment>", "Target environment", "dev")
    .action(
      wrapHandler("teamsapp publish", async (_opts, _cmd) => {
        console.log("Publishing Teams app...");
      })
    );

  // atk teamsapp update (stub)
  teamsapp
    .command("update")
    .description("Update the Teams app registration")
    .option("-e, --env <environment>", "Target environment", "dev")
    .action(
      wrapHandler("teamsapp update", async (_opts, _cmd) => {
        console.log("Updating Teams app...");
      })
    );

  // atk teamsapp doctor (stub)
  teamsapp
    .command("doctor")
    .description("Check prerequisites and troubleshoot project issues")
    .action(
      wrapHandler("teamsapp doctor", async (_opts, _cmd) => {
        console.log("Running diagnostics...");
      })
    );
}

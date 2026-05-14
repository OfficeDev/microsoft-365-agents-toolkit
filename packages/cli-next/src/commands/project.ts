// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Command } from "commander";
import { templateRegistry, registerBuiltinTemplates } from "@microsoft/teamsfx-core-next";
import { wrapHandler, wrapHandlerWithContext, renderPostActions } from "../handler";
import { buildNewCommands } from "./factory";
import { provisionAction, deployAction, publishAction } from "../actions/lifecycle";

/**
 * Project lifecycle commands: new, provision, deploy, publish, preview, upgrade
 */
export function createProjectCommands(program: Command): void {
  // Ensure built-in templates are loaded
  registerBuiltinTemplates();

  // atk new — registry-driven subcommands
  const newCmd = program
    .command("new")
    .description("Create a new Microsoft 365 app or agent project");

  // Build subcommands from the template registry (atk new da, atk new bot, etc.)
  buildNewCommands(newCmd, templateRegistry);

  // atk new sample (keep as stub — not backed by template registry)
  newCmd
    .command("sample")
    .description("Create a project from a sample template")
    .option("-n, --name <name>", "Sample name")
    .option("-f, --folder <folder>", "Target folder", ".")
    .action(
      wrapHandler("new sample", async (_opts, _cmd) => {
        console.log("Creating from sample...");
      })
    );

  // atk provision
  program
    .command("provision")
    .description("Provision cloud resources for the project")
    .option("-e, --env <environment>", "Target environment name", "dev")
    .option("--project-folder <folder>", "Path to the project folder")
    .option("--skip-consent", "Skip the confirmation prompt", false)
    .action(
      wrapHandlerWithContext("provision", async (ctx, opts) => {
        const result = await provisionAction(ctx, {
          projectPath: ctx.projectPath ?? process.cwd(),
          envName: (opts.env as string) ?? "dev",
          skipConsent: (opts.skipConsent as boolean) ?? false,
        });
        renderPostActions(result.postActions);
      })
    );

  // atk deploy
  program
    .command("deploy")
    .description("Deploy the project to cloud")
    .option("-e, --env <environment>", "Target environment name", "dev")
    .option("--project-folder <folder>", "Path to the project folder")
    .option("--skip-consent", "Skip the confirmation prompt", false)
    .action(
      wrapHandlerWithContext("deploy", async (ctx, opts) => {
        const result = await deployAction(ctx, {
          projectPath: ctx.projectPath ?? process.cwd(),
          envName: (opts.env as string) ?? "dev",
          skipConsent: (opts.skipConsent as boolean) ?? false,
        });
        renderPostActions(result.postActions);
      })
    );

  // atk publish
  program
    .command("publish")
    .description("Publish the app to your organization's Teams app catalog")
    .option("-e, --env <environment>", "Target environment name", "dev")
    .option("--project-folder <folder>", "Path to the project folder")
    .action(
      wrapHandlerWithContext("publish", async (ctx, opts) => {
        const result = await publishAction(ctx, {
          projectPath: ctx.projectPath ?? process.cwd(),
          envName: (opts.env as string) ?? "dev",
        });
        renderPostActions(result.postActions);
      })
    );

  // atk preview (stub — not yet backed by core-next)
  program
    .command("preview")
    .description("Run and preview the app locally")
    .option("-e, --env <environment>", "Target environment name", "local")
    .option("--browser <browser>", "Browser to open: default, chrome, edge, firefox", "default")
    .option("--m365-host <host>", "M365 host to preview in: Teams, Outlook, Office", "Teams")
    .action(
      wrapHandler("preview", async (_opts, _cmd) => {
        console.log("Starting local preview...");
      })
    );

  // atk upgrade (stub — not yet backed by core-next)
  program
    .command("upgrade")
    .description("Upgrade the project to the latest version")
    .option("--force", "Force upgrade without confirmation", false)
    .action(
      wrapHandler("upgrade", async (_opts, _cmd) => {
        console.log("Upgrading project...");
      })
    );
}

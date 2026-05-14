// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Command } from "commander";
import { wrapHandler } from "../handler";
import { envListAction, envAddAction, envResetAction } from "../actions/environment";
import { colorize, TextType } from "../output";

/**
 * Environment commands: env add, env list, env reset
 */
export function createEnvCommands(program: Command): void {
  const env = program.command("env").description("Manage project environments");

  // atk env add
  env
    .command("add")
    .description("Create a new environment for the project")
    .argument("<name>", "Name of the new environment")
    .option("--copy-from <env>", "Copy configuration from existing environment")
    .option("--project-folder <folder>", "Path to the project folder")
    .action(
      wrapHandler("env add", async (opts, cmd) => {
        const name = cmd.args[0];
        const projectPath = (opts.projectFolder as string) ?? process.cwd();
        await envAddAction(projectPath, name, opts.copyFrom as string | undefined);
        console.log(colorize(`Environment "${name}" created.`, TextType.Success));
      })
    );

  // atk env list
  env
    .command("list")
    .description("List all environments for the project")
    .option("--project-folder <folder>", "Path to the project folder")
    .action(
      wrapHandler("env list", async (opts) => {
        const projectPath = (opts.projectFolder as string) ?? process.cwd();
        const envs = await envListAction(projectPath);
        if (envs.length === 0) {
          console.log("No environments found.");
        } else {
          for (const name of envs) {
            console.log(`  ${name}`);
          }
        }
      })
    );

  // atk env reset
  env
    .command("reset")
    .description("Reset an environment's configuration")
    .argument("<name>", "Name of the environment to reset")
    .option("--project-folder <folder>", "Path to the project folder")
    .action(
      wrapHandler("env reset", async (opts, cmd) => {
        const name = cmd.args[0];
        const projectPath = (opts.projectFolder as string) ?? process.cwd();
        await envResetAction(projectPath, name);
        console.log(colorize(`Environment "${name}" has been reset.`, TextType.Success));
      })
    );
}

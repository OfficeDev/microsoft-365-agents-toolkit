// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Command } from "commander";
import { createProjectCommands } from "./project";
import { createAccountCommands } from "./account";
import { createEnvCommands } from "./env";
import { createTeamsappCommands } from "./teamsapp";
import { createAddCommands } from "./add";
import { createListCommands } from "./list";
import { createM365Commands } from "./m365";
import { createPermissionCommands } from "./permission";
import { createEntraAppCommands } from "./entraApp";
import { createRegenerateCommands } from "./regenerate";
import { createMiscCommands } from "./misc";

/**
 * Build the full Commander.js program with all command groups.
 */
export function buildProgram(binName: string): Command {
  const program = new Command(binName);

  program
    .description("Microsoft 365 Agents Toolkit CLI")
    .version("4.0.0", "-v, --version", "Output the current version")
    .option("--output <format>", "Output format: text | json", "text")
    .option("--non-interactive", "Run in non-interactive mode (no prompts)")
    .option("--debug", "Enable debug logging")
    .configureHelp({
      sortSubcommands: true,
      sortOptions: true,
    });

  // Register all command groups
  createProjectCommands(program);
  createAccountCommands(program);
  createEnvCommands(program);
  createTeamsappCommands(program);
  createAddCommands(program);
  createListCommands(program);
  createM365Commands(program);
  createPermissionCommands(program);
  createEntraAppCommands(program);
  createRegenerateCommands(program);
  createMiscCommands(program);

  return program;
}

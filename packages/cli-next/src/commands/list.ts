// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Command } from "commander";
import { templateRegistry, registerBuiltinTemplates } from "@microsoft/teamsfx-core-next";
import { wrapHandler } from "../handler";
import { listTemplatesAction } from "../actions/listTemplates";

/**
 * List commands: list, list samples, list templates
 */
export function createListCommands(program: Command): void {
  const list = program.command("list").description("List project info, samples, or templates");

  // atk list templates
  list
    .command("templates")
    .description("List all available project templates")
    .action(
      wrapHandler("list templates", async () => {
        registerBuiltinTemplates();
        const rows = listTemplatesAction(templateRegistry);
        if (rows.length === 0) {
          console.log("No templates registered.");
          return;
        }
        // Simple table output
        const header = `${"ID".padEnd(35)} ${"Name".padEnd(30)} ${"Category".padEnd(20)} Languages`;
        console.log(header);
        console.log("-".repeat(header.length));
        for (const row of rows) {
          console.log(
            `${row.id.padEnd(35)} ${row.name.padEnd(30)} ${row.category.padEnd(20)} ${row.languages}`
          );
        }
      })
    );

  // atk list samples
  list
    .command("samples")
    .description("List all available sample projects")
    .option("--tag <tag>", "Filter samples by tag")
    .action(
      wrapHandler("list samples", async (_opts, _cmd) => {
        console.log("Listing samples...");
      })
    );
}

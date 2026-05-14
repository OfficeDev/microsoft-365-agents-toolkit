// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Command } from "commander";
import { wrapHandler } from "../handler";

/**
 * Regenerate commands: regenerate, regenerate action
 */
export function createRegenerateCommands(program: Command): void {
  const regenerate = program.command("regenerate").description("Regenerate project resources");

  // atk regenerate action
  regenerate
    .command("action")
    .description("Regenerate an action from updated API specification")
    .option("--api-spec-path <path>", "Path to the updated OpenAPI specification")
    .action(
      wrapHandler("regenerate action", async (_opts, _cmd) => {
        console.log("Regenerating action...");
      })
    );
}

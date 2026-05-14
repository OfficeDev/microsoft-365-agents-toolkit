// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as path from "path";
import * as fs from "fs";
import { Command } from "commander";
import { buildProgram } from "./commands";
import { cliTelemetry } from "./telemetry";
import { logger } from "./logger";
import { cliUI } from "./ui";
import { colorize, TextType } from "./output";

/**
 * CLI entry point. Called from cli.js / cliold.js.
 */
export async function start(): Promise<void> {
  const binName = process.env.TEAMSFX_CLI_BIN_NAME ?? "atk";

  // Initialise telemetry — reads aiKey from own package.json.
  // If the key is absent (local dev), telemetry stays a no-op.
  try {
    const pkgPath = path.resolve(__dirname, "..", "package.json");
    const pkg = JSON.parse(fs.readFileSync(pkgPath, "utf-8"));
    cliTelemetry.init(pkg.aiKey ?? "", pkg.version ?? "0.0.0");
  } catch {
    // If reading package.json fails, telemetry just stays inactive
  }

  cliTelemetry.addSharedProperty("binName", binName);

  if (binName === "teamsapp") {
    await logger.warning(
      'Deprecation Warning: The CLI command "teamsapp" is renamed to "atk". ' +
        'Please use "atk" instead. "teamsapp" will be removed in a future release.'
    );
  }

  const program = buildProgram(binName);

  // Wire global options after parse
  program.hook("preAction", (thisCommand: Command) => {
    const globalOpts = thisCommand.opts();

    // Non-interactive mode
    if (globalOpts.nonInteractive || process.env.CI_ENABLED === "true") {
      cliUI.interactive = false;
    }

    // Debug logging
    if (globalOpts.debug) {
      logger.setLogLevel(1); // LogLevel.Debug
    }
  });

  // Handle unknown commands gracefully
  program.showSuggestionAfterError(true);

  try {
    await program.parseAsync(process.argv);
  } catch (error: unknown) {
    if (error instanceof Error) {
      console.error(colorize(`Fatal: ${error.message}`, TextType.Error));
    }
    process.exitCode = 1;
  }
}

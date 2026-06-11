// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import fs from "fs-extra";
import * as path from "path";
import * as constants from "./constants";
import cliTelemetry from "./telemetry/cliTelemetry";
import { TelemetryProperty } from "./telemetry/cliTelemetryEvents";

export function initTelemetryReporter(): void {
  const { CliTelemetryReporter } =
    require("./commonlib/telemetry") as typeof import("./commonlib/telemetry");
  const cliPackage = JSON.parse(fs.readFileSync(path.join(__dirname, "/../package.json"), "utf8"));
  const reporter = new CliTelemetryReporter(
    cliPackage.aiKey,
    constants.cliTelemetryPrefix,
    cliPackage.version
  );
  cliTelemetry.reporter = reporter;
}

/**
 * Starts the CLI process.
 */
export async function start(): Promise<void> {
  const { logger } = require("./commonlib/logger") as typeof import("./commonlib/logger");
  const { start: startNewUX } = require("./commands/index") as typeof import("./commands/index");

  initTelemetryReporter();
  const binName = process.env.TEAMSFX_CLI_BIN_NAME as string;
  if (binName === "teamsapp") {
    logger.warning(
      `Deprecation Warning: The CLI command "teamsapp" is renamed to "atk". The old command name will be retired soon. Please switch to the new command and update your workflows accordingly.`
    );
  }
  cliTelemetry.reporter?.addSharedProperty(TelemetryProperty.BinName, binName); // trigger binary name for telemetry
  return startNewUX(binName);
}

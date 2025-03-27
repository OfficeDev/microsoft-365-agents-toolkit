// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import fs from "fs-extra";
import * as path from "path";
import { start as startNewUX } from "./commands/index";
import { CliTelemetryReporter } from "./commonlib/telemetry";
import "./console/screen";
import * as constants from "./constants";
import cliTelemetry from "./telemetry/cliTelemetry";
import { TelemetryProperty } from "./telemetry/cliTelemetryEvents";
import { logger } from "./commonlib/logger";

export function initTelemetryReporter(): void {
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
  initTelemetryReporter();
  const binName = process.env.TEAMSFX_CLI_BIN_NAME as string;
  if (binName === "teamsapp") {
    logger.warning(
      `Deprecation Warning: The CLI package "@microsoft/teamsapp-cli" and its "teamsapp" command are renamed to "@microsoft/m365agentstoolkit-cli" and "m365agents" respectively. The old command "teamsapp" will be retired soon. Please switch to the new command "m365agents" and update your workflows accordingly.`
    );
  }
  cliTelemetry.reporter?.addSharedProperty(TelemetryProperty.BinName, binName); // trigger binary name for telemetry
  return startNewUX(binName);
}

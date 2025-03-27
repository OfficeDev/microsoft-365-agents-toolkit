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
  cliTelemetry.reporter?.addSharedProperty(TelemetryProperty.BinName, binName); // trigger binary name for telemetry
  return startNewUX(binName);
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Stage } from "@microsoft/teamsfx-api";
import chalk from "chalk";

export const cliSource = "TeamsfxCLI";
export const cliName = "teamsfx";
export const cliTelemetryPrefix = "teamsfx-cli";
export const teamsAppFileName = "teamsapp.yml";

export enum CLILogLevel {
  error = 0,
  verbose,
  debug,
}

export const AddFeatureFunc = {
  namespace: "fx-solution-azure",
  method: Stage.addFeature,
};

export const FooterText = `For more information about the Microsoft 365 Agents Toolkit: ${chalk.cyanBright(
  "https://aka.ms/teamsfx-toolkit-cli"
)}.`;

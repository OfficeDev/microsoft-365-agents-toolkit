// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand } from "@microsoft/teamsfx-api";
import { commands } from "../../resource";
import { exportOpenPluginCommand } from "./exportOpenPlugin";

export function exportCommand(): CLICommand {
  return {
    name: "export",
    description: commands.export.description,
    commands: [exportOpenPluginCommand],
  };
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand } from "@microsoft/teamsfx-api";
import { commands } from "../../resource";
import { importOpenPluginCommand } from "./importOpenPlugin";

export function importCommand(): CLICommand {
  return {
    name: "import",
    description: commands.import.description,
    commands: [importOpenPluginCommand],
  };
}

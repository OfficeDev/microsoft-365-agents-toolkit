// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand } from "@microsoft/teamsfx-api";
import { commands } from "../../resource";
import { regeneratePluginCommand } from "./regeneratePlugin";

const adjustCommands = (): CLICommand[] => {
  return [regeneratePluginCommand];
};
export function regenerateCommand(): CLICommand {
  return {
    name: "regenerate",
    description: commands.regenerate.description,
    commands: adjustCommands(),
  };
}

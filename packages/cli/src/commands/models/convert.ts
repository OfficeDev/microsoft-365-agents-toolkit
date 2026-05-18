// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand } from "@microsoft/teamsfx-api";
import { commands } from "../../resource";
import { convertOpenPluginCommand } from "./convertOpenPlugin";

export function convertCommand(): CLICommand {
  return {
    name: "convert",
    description: commands.convert.description,
    commands: [convertOpenPluginCommand],
  };
}

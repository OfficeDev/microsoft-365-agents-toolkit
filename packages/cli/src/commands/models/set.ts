// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand } from "@microsoft/teamsfx-api";
import { commands } from "../../resource";
import { setSensitivityLabelCommand } from "./setSensitivityLabel";

const adjustCommands = (): CLICommand[] => {
  return [setSensitivityLabelCommand];
};

export function setCommand(): CLICommand {
  return {
    name: "set",
    description: commands.set.description,
    commands: adjustCommands(),
  };
}

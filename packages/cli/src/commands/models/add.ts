// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand } from "@microsoft/teamsfx-api";
import { commands } from "../../resource";
import { addPluginCommand } from "./addPlugin";
import { addAuthConfigCommand } from "./addAuthConfig";
import { addCapabilityCommand } from "./addCapability";
import { addSkillCommand } from "./addSkill";

const adjustCommands = (): CLICommand[] => {
  return [addPluginCommand, addAuthConfigCommand, addCapabilityCommand, addSkillCommand];
};
export function addCommand(): CLICommand {
  return {
    name: "add",
    description: commands.add.description,
    commands: adjustCommands(),
  };
}

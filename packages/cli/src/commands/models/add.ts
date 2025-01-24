// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand } from "@microsoft/teamsfx-api";
import { commands } from "../../resource";
import { addSPFxWebpartCommand } from "./addSPFxWebpart";
import { addPluginCommand } from "./addPlugin";
import { addKnowledgeCommand } from "./addKnowledge";
import { addAuthConfigCommand } from "./addAuthConfig";

const adjustCommands = (): CLICommand[] => {
  return [addSPFxWebpartCommand, addPluginCommand, addAuthConfigCommand, addKnowledgeCommand];
};
export function addCommand(): CLICommand {
  return {
    name: "add",
    description: commands.add.description,
    commands: adjustCommands(),
  };
}

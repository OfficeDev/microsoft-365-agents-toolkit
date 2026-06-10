// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand } from "@microsoft/teamsfx-api";
import { featureFlagManager, FeatureFlags } from "@microsoft/teamsfx-core";
import { commands } from "../../resource";
import { addSPFxWebpartCommand } from "./addSPFxWebpart";
import { addPluginCommand } from "./addPlugin";
import { addAuthConfigCommand } from "./addAuthConfig";
import { addCapabilityCommand } from "./addCapability";
import { addSkillCommand } from "./addSkill";

const adjustCommands = (): CLICommand[] => {
  return [
    addSPFxWebpartCommand,
    addPluginCommand,
    addAuthConfigCommand,
    addCapabilityCommand,
    ...(featureFlagManager.getBooleanValue(FeatureFlags.AgentSkillsManifest)
      ? [addSkillCommand]
      : []),
  ];
};
export function addCommand(): CLICommand {
  return {
    name: "add",
    description: commands.add.description,
    commands: adjustCommands(),
  };
}

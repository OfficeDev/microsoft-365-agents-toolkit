// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand } from "@microsoft/teamsfx-api";
import { commands } from "../../resource";
import { featureFlagManager, FeatureFlags } from "@microsoft/teamsfx-core";
import { setSensitivityLabelCommand } from "./setSensitivityLabel";

const adjustCommands = (): CLICommand[] => {
  if (featureFlagManager.getBooleanValue(FeatureFlags.SensitivityLabelEnabled)) {
    return [setSensitivityLabelCommand];
  }
  return [];
};

export function setCommand(): CLICommand {
  return {
    name: "set",
    description: commands.set.description,
    commands: adjustCommands(),
  };
}

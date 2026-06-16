// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand } from "@microsoft/teamsfx-api";
import { AddSkillInputs, AddSkillOptions } from "@microsoft/teamsfx-core";
import { getFxCore } from "../../activate";
import { commands } from "../../resource";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";
import { ProjectFolderOption } from "../common";

export const addSkillCommand: CLICommand = {
  name: "skill",
  description: commands["add.skill"].description,
  options: [...AddSkillOptions, ProjectFolderOption],
  telemetry: {
    event: TelemetryEvent.AddSkill,
  },
  handler: async (ctx) => {
    const inputs = ctx.optionValues as AddSkillInputs;
    const core = getFxCore();
    const res = await core.addSkill(inputs);
    return res;
  },
};

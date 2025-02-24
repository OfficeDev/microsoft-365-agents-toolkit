// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand } from "@microsoft/teamsfx-api";
import { AddKnowledgeInputs, AddKnowledgeOptions } from "@microsoft/teamsfx-core";
import { getFxCore } from "../../activate";
import { commands } from "../../resource";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";

export const addKnowledgeCommand: CLICommand = {
  name: "knowledge",
  description: commands["add.knowledge"].description,
  options: [...AddKnowledgeOptions],
  telemetry: {
    event: TelemetryEvent.AddKnowledge,
  },
  handler: async (ctx) => {
    const inputs = ctx.optionValues as AddKnowledgeInputs;
    const core = getFxCore();
    const res = await core.addKnowledge(inputs);
    return res;
  },
};

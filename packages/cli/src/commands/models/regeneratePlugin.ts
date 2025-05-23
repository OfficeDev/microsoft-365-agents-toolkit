// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand } from "@microsoft/teamsfx-api";
import { RegeneratePluginInputs, RegeneratePluginOptions } from "@microsoft/teamsfx-core";
import { getFxCore } from "../../activate";
import { commands } from "../../resource";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";
import { ProjectFolderOption } from "../common";

export const regeneratePluginCommand: CLICommand = {
  name: "action",
  description: commands["regenerate.action"].description,
  options: [...RegeneratePluginOptions, ProjectFolderOption],
  telemetry: {
    event: TelemetryEvent.RegneratePlugin,
  },
  handler: async (ctx) => {
    const inputs = ctx.optionValues as RegeneratePluginInputs;
    const core = getFxCore();
    const res = await core.regeneratePlugin(inputs);
    return res;
  },
};

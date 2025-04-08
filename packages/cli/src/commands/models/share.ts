// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand, CLIContext, InputsWithProjectPath } from "@microsoft/teamsfx-api";
import { getFxCore } from "../../activate";
import { commands } from "../../resource";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";
import { EnvOption, IgnoreLoadEnvOption, ProjectFolderOption } from "../common";
import { ShareOptions } from "@microsoft/teamsfx-core";
import { shareRemoveCommand } from "./shareRemove";

export const shareCommand: CLICommand = {
  name: "share",
  description: commands.share.description,
  options: [EnvOption, ProjectFolderOption, IgnoreLoadEnvOption, ...ShareOptions],
  telemetry: {
    event: TelemetryEvent.Share,
  },
  handler: async (ctx: CLIContext) => {
    const inputs = ctx.optionValues as InputsWithProjectPath;
    const core = getFxCore();
    const res = await core.shareApplication(inputs);
    return res;
  },
  commands: [shareRemoveCommand],
};

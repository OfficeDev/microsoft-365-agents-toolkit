// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand, CLIContext, InputsWithProjectPath } from "@microsoft/teamsfx-api";
import { ShareOptions } from "@microsoft/teamsfx-core";
import { getFxCore } from "../../activate";
import { commands } from "../../resource";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";
import { EnvOption, IgnoreLoadEnvOption, ProjectFolderOption } from "../common";
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
  examples: [
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} share`,
      description: "Share under current project folder in interactive mode",
    },
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} share --scope tenant -i false`,
      description: "Share the agent to all tenant users",
    },
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} share --scope users --email 'a@example.com,b@example.com' -i false`,
      description: "Share the agent to specific users or groups",
    },
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} share --scope owner --email 'a@example.com,b@example.com' -i false`,
      description: "Share the ownership of agent with users",
    },
  ],
  commands: [shareRemoveCommand],
};

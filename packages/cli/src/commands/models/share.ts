// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  CLICommand,
  CLICommandOption,
  CLIContext,
  InputsWithProjectPath,
} from "@microsoft/teamsfx-api";
import { QuestionNames, ShareOperationOption, ShareOptions } from "@microsoft/teamsfx-core";
import { getFxCore } from "../../activate";
import { commands } from "../../resource";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";
import { EnvOption, IgnoreLoadEnvOption, ProjectFolderOption } from "../common";
import { shareRemoveCommand } from "./shareRemove";

const shareOptions = ShareOptions.filter((option: CLICommandOption) => {
  return ["scope", "email"].includes(option.name);
});

export const shareCommand: CLICommand = {
  name: "share",
  description: commands.share.description,
  options: [EnvOption, ProjectFolderOption, IgnoreLoadEnvOption, ...shareOptions],
  telemetry: {
    event: TelemetryEvent.Share,
  },
  handler: async (ctx: CLIContext) => {
    const inputs = ctx.optionValues as InputsWithProjectPath;
    inputs[QuestionNames.ShareOperation] = ShareOperationOption.ShareWithUsers;
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
      description: "Share the agent with all tenant users",
    },
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} share --scope users --email 'a@example.com,b@example.com' -i false`,
      description: "Share the agent with specific users",
    },
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} share --scope owners --email 'a@example.com,b@example.com' -i false`,
      description: "Share the ownership of agent with selected users",
    },
  ],
  commands: [shareRemoveCommand],
};

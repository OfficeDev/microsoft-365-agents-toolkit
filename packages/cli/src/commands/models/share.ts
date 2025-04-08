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
  examples: [
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} share`,
      description: "Share under current project folder in interactive mode",
    },
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} share --option share-app -i false`,
      description: "Share the app under current project folder",
    },
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} share --option share-to-users --users 'a@example.com,b@example.com' -i false`,
      description: "Share owner access to users",
    },
  ],
  commands: [shareRemoveCommand],
};

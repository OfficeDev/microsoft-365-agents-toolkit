// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand, InputsWithProjectPath } from "@microsoft/teamsfx-api";
import { getFxCore } from "../../activate";
import { commands } from "../../resource";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";
import { EnvOption, IgnoreLoadEnvOption, ProjectFolderOption } from "../common";
import { removeSharedAccessOptions } from "@microsoft/teamsfx-core";

export const shareRemoveCommand: CLICommand = {
  name: "remove",
  description: commands["share.remove"].description,
  options: [EnvOption, ProjectFolderOption, IgnoreLoadEnvOption, ...removeSharedAccessOptions],
  telemetry: {
    event: TelemetryEvent.ShareRemove,
  },
  examples: [
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} share remove`,
      description: "Remove shared owner access under current project folder in interactive mode",
    },
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} share remove --users 'a@example.com,b@example.com' -i false`,
      description: "Remove shared owner access from users",
    },
  ],
  handler: async (ctx) => {
    const inputs = ctx.optionValues as InputsWithProjectPath;
    const core = getFxCore();
    const res = await core.removeSharedAccess(inputs);
    return res;
  },
};

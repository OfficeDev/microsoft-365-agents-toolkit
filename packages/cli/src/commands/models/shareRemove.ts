// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand, CLICommandOption, InputsWithProjectPath } from "@microsoft/teamsfx-api";
// import { removeSharedAccessOptions } from "@microsoft/teamsfx-core";
import { QuestionNames, ShareOperationOption } from "@microsoft/teamsfx-core";
import { getFxCore } from "../../activate";
import { commands } from "../../resource";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";
import { EnvOption, IgnoreLoadEnvOption, ProjectFolderOption } from "../common";

// Use customized option to temporarily keep owner access removal before moving to collaborator
const removeSharedAccessOptions: CLICommandOption[] = [
  {
    name: "users",
    type: "string",
    description: "Remove access for selected user(s)",
    required: false,
    skipValidation: true,
  },
  {
    name: "owners",
    type: "array",
    description: "Select owners to remove access",
    required: false,
    skipValidation: true,
  },
];

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
      description: "Remove shared access from users",
    },
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} share remove --owners 'a@example.com,b@example.com' -i false`,
      description: "Remove shared ownership from users",
    },
  ],
  handler: async (ctx) => {
    const inputs = ctx.optionValues as InputsWithProjectPath;
    const core = getFxCore();
    if (inputs["users"]) {
      inputs[QuestionNames.ShareOperation] = ShareOperationOption.RemoveShareAccessFromUsers;
      inputs[QuestionNames.UserEmail] = inputs["users"];
      return await core.shareApplication(inputs);
    } else {
      inputs["users"] = inputs["owners"];
      return await core.removeSharedAccess(inputs);
    }
  },
};

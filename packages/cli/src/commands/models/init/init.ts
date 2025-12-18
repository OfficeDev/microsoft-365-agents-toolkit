// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand, err, Inputs, ok } from "@microsoft/teamsfx-api";
import { getFxCore } from "../../../activate";
import { commands } from "../../../resource";
import { TelemetryEvent } from "../../../telemetry/cliTelemetryEvents";
import { ProjectFolderOptionWithoutValidation, TeamsAppManifestFileOption } from "../../common";
import {
  localDebugOption,
  playgroundOption,
  programmingLanguageOption,
  remoteDeployOption,
} from "./initOption";

export const initCommand: CLICommand = {
  name: "init",
  description: commands.init.description,
  options: [
    playgroundOption,
    localDebugOption,
    remoteDeployOption,
    programmingLanguageOption,
    { ...TeamsAppManifestFileOption, required: true },
    ProjectFolderOptionWithoutValidation,
  ],
  defaultInteractiveOption: false,
  telemetry: {
    event: TelemetryEvent.InitProject,
  },
  examples: [
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} init`,
      description:
        "Initialize current project as a Microsoft 365 Agents Toolkit project, local debug is enabled by default",
    },
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} init --remote true`,
      description:
        "Initialize current project as a Microsoft 365 Agents Toolkit project including remote deployment configuration",
    },
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} init --playground true --local false`,
      description:
        "Initialize current project as a Microsoft 365 Agents Toolkit project with Playground debug option only",
    },
  ],
  handler: async (ctx) => {
    const inputs = ctx.optionValues;
    const core = getFxCore();
    const result = await core.generateConfigFiles(inputs as Inputs);
    if (result.isErr()) {
      return err(result.error);
    }
    return ok(undefined);
  },
};

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
    event: TelemetryEvent.GenerateConfig,
  },
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

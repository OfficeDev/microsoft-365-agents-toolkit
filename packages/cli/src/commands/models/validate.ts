// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand, Result, err, ok } from "@microsoft/teamsfx-api";
import { ValidateTeamsAppInputs, ValidateTeamsAppOptions } from "@microsoft/teamsfx-core";
import { getFxCore } from "../../activate";
import { ArgumentConflictError, MissingRequiredOptionError } from "../../error";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";
import { EnvOption, ProjectFolderOption } from "../common";
import * as path from "path";
import { commands } from "../../resource";

export const validateCommand: CLICommand = {
  name: "validate",
  description: commands.validate.description,
  options: [...ValidateTeamsAppOptions, EnvOption, ProjectFolderOption],
  telemetry: {
    event: TelemetryEvent.ValidateManifest,
  },
  examples: [
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} validate --app-package-file ./appPackage/build/appPackage.dev.zip`,
      description: "Validate the Microsoft 365 App package.",
    },
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} validate --teams-manifest-file ./appPackage/manifest.json --env dev`,
      description: "Validate the Microsoft 365 App manifest using its schema.",
    },
  ],
  defaultInteractiveOption: false,
  handler: async (ctx) => {
    const inputs = ctx.optionValues as ValidateTeamsAppInputs;
    if (!ctx.globalOptionValues.interactive) {
      const res = validateInputs(ctx.command.fullName, inputs);
      if (res.isErr()) {
        return err(res.error);
      }
    }
    const core = getFxCore();
    const res = await core.validateApplication(inputs);
    if (res.isErr()) {
      return err(res.error);
    }
    return ok(undefined);
  },
};

function validateInputs(
  fullName: string,
  inputs: ValidateTeamsAppInputs
): Result<
  undefined,
  ArgumentConflictError | MissingRequiredOptionError | MissingRequiredOptionError
> {
  if (inputs["manifest-path"] && inputs["app-package-file-path"]) {
    const error = new ArgumentConflictError(
      fullName,
      "teams-manifest-file",
      "app-package-file-path"
    );
    return err(error);
  } else if (!inputs["manifest-path"] && !inputs["app-package-file-path"]) {
    inputs["manifest-path"] = path.join(
      path.resolve(inputs.projectPath! || "./"),
      "./appPackage/manifest.json"
    );
  }
  if (!inputs["app-package-file-path"] && !inputs.env) {
    return err(new MissingRequiredOptionError(fullName, "--env"));
  }
  return ok(undefined);
}

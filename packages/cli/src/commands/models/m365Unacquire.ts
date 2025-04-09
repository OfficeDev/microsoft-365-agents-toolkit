// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand, err, ok } from "@microsoft/teamsfx-api";
import { UninstallInputs, QuestionNames } from "@microsoft/teamsfx-core";
import { logger } from "../../commonlib/logger";
import { MissingRequiredOptionError } from "../../error";
import { commands } from "../../resource";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";
import { m365utils, sideloadingServiceEndpoint } from "./m365Sideloading";
import { getFxCore } from "../../activate";

export const m365UnacquireCommand: CLICommand = {
  name: "uninstall",
  aliases: ["unacquire"],
  description: commands.uninstall.description,
  options: [
    {
      name: QuestionNames.UninstallMode,
      description: commands.uninstall.options["mode"],
      type: "string",
    },
    {
      name: QuestionNames.TitleId,
      description: commands.uninstall.options["title-id"],
      type: "string",
    },
    {
      name: QuestionNames.ManifestId,
      description: commands.uninstall.options["manifest-id"],
      type: "string",
    },
    {
      name: QuestionNames.Env,
      description: commands.uninstall.options["env"],
      type: "string",
    },
    {
      name: "folder",
      questionName: QuestionNames.ProjectPath,
      description: commands.uninstall.options["folder"],
      type: "string",
    },
    {
      name: QuestionNames.UninstallOptions,
      description: commands.uninstall.options["options"],
      type: "array",
    },
  ],
  examples: [
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} uninstall -i false --mode title-id --title-id U_xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`,
      description: "Remove the acquired Microsoft 365 Application using Title ID",
    },
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} uninstall -i false --mode manifest-id --manifest-id xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx --options 'm365-app,app-registration,bot-framework-registration'`,
      description: "Remove the acquired Microsoft 365 Application using Manifest ID",
    },
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} uninstall -i false --mode env --env xxx --options 'm365-app,app-registration,bot-framework-registration' --folder ./myapp`,
      description:
        "Remove the acquired Microsoft 365 Application using environment in Microsoft 365 Agents Toolkit generated project",
    },
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} uninstall`,
      description: "Uninstall in interactive mode",
    },
  ],
  telemetry: {
    event: TelemetryEvent.Uninstall,
  },
  defaultInteractiveOption: true,
  handler: async (ctx) => {
    const inputs = ctx.optionValues as UninstallInputs;
    const core = getFxCore();
    const res = await core.uninstall(inputs);
    if (res.isErr()) {
      return err(res.error);
    }
    return ok(undefined);
  },
};

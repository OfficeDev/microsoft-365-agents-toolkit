// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand } from "@microsoft/teamsfx-api";
import { commands, strings } from "../../resource";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";
import { ProjectFolderOption } from "../common";
import { getFxCore } from "../../activate";
import { AddAuthActionInputs, AddAuthActionOptions } from "@microsoft/teamsfx-core";
import { logger } from "../../commonlib/logger";

export const addAuthConfigCommand: CLICommand = {
  name: "auth-config",
  description: commands["add.auth-config"].description,
  options: [...AddAuthActionOptions, ProjectFolderOption],
  telemetry: {
    event: TelemetryEvent.AddAuthAction,
  },
  handler: async (ctx) => {
    const inputs = ctx.optionValues as AddAuthActionInputs;
    const core = getFxCore();
    const res = await core.addAuthAction(inputs);
    logger.info(strings.command.add["auth-config"].notification);
    return res;
  },
};

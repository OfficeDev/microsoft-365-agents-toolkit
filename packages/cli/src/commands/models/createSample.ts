// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CLICommand, CLIContext, err, ok } from "@microsoft/teamsfx-api";
import {
  CreateSampleProjectArguments,
  CreateSampleProjectInputs,
  CreateSampleProjectOptions,
} from "@microsoft/teamsfx-core";
import chalk from "chalk";
import { assign } from "lodash";
import * as path from "path";
import * as uuid from "uuid";
import * as activate from "../../activate";
import { logger } from "../../commonlib/logger";
import { commands } from "../../resource";
import { TelemetryEvent, TelemetryProperty } from "../../telemetry/cliTelemetryEvents";

export const createSampleCommand: CLICommand = {
  name: "sample",
  description: commands["create.sample"].description,
  arguments: CreateSampleProjectArguments,
  options: CreateSampleProjectOptions,
  telemetry: {
    event: TelemetryEvent.DownloadSample,
  },
  handler: async (ctx: CLIContext) => {
    const inputs = ctx.optionValues as CreateSampleProjectInputs;
    inputs.projectId = inputs.projectId ?? uuid.v4();
    const core = activate.getFxCore();
    const res = await core.createSampleProject(inputs);
    assign(ctx.telemetryProperties, {
      [TelemetryProperty.NewProjectId]: inputs.projectId,
      [TelemetryProperty.SampleName]: inputs.samples,
    });
    if (res.isErr()) {
      return err(res.error);
    }
    logger.info(
      `Sample project '${chalk.cyan(inputs.samples)}' downloaded at: ${chalk.cyan(
        path.resolve(res.value.projectPath)
      )}`
    );
    return ok(undefined);
  },
};

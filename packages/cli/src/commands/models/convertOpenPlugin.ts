// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand, err, ok } from "@microsoft/teamsfx-api";
import {
  ConvertOpenPluginInputs,
  ConvertOpenPluginOptions,
} from "@microsoft/teamsfx-core";
import { getFxCore } from "../../activate";
import { logger } from "../../commonlib/logger";
import { commands } from "../../resource";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";

export const convertOpenPluginCommand: CLICommand = {
  name: "openplugin",
  description: commands["convert.openplugin"].description,
  options: [...ConvertOpenPluginOptions],
  telemetry: {
    event: TelemetryEvent.ConvertOpenPlugin,
  },
  defaultInteractiveOption: false,
  handler: async (ctx) => {
    const inputs = ctx.optionValues as ConvertOpenPluginInputs;
    const core = getFxCore();
    const res = await core.convertOpenPlugin(inputs);
    if (res.isErr()) {
      return err(res.error);
    }
    logger.info(`Project created at: ${res.value.projectPath}`);
    for (const warning of res.value.warnings ?? []) {
      logger.warning(warning.content);
    }
    return ok(undefined);
  },
};

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand, err, ok } from "@microsoft/teamsfx-api";
import { ExportOpenPluginInputs, ExportOpenPluginOptions } from "@microsoft/teamsfx-core";
import { getFxCore } from "../../activate";
import { logger } from "../../commonlib/logger";
import { commands } from "../../resource";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";

export const exportOpenPluginCommand: CLICommand = {
  name: "openplugin",
  description: commands["export.openplugin"].description,
  options: [...ExportOpenPluginOptions],
  telemetry: {
    event: TelemetryEvent.ExportOpenPlugin,
  },
  defaultInteractiveOption: false,
  handler: async (ctx) => {
    const inputs = ctx.optionValues as ExportOpenPluginInputs;
    const core = getFxCore();
    const res = await core.exportOpenPlugin(inputs);
    if (res.isErr()) {
      return err(res.error);
    }
    logger.info(`Open Plugin written to: ${res.value.outputPath}`);
    for (const warning of res.value.warnings ?? []) {
      logger.warning(warning.content);
    }
    return ok(undefined);
  },
};

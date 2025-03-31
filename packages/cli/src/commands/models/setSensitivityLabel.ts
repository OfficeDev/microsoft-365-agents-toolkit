// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand } from "@microsoft/teamsfx-api";
import { SetSensitivityLabelInputs, SetSensitivityLabelOptions } from "@microsoft/teamsfx-core";
import { getFxCore } from "../../activate";
import { commands } from "../../resource";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";
import { ProjectFolderOptionWithoutValidation } from "../common";

export const setSensitivityLabelCommand: CLICommand = {
  name: "sensitivityLabel",
  description: commands["set.sensitivityLabel"].description,
  options: [...SetSensitivityLabelOptions, ProjectFolderOptionWithoutValidation],
  telemetry: {
    event: TelemetryEvent.SetSensitivityLabel,
  },
  handler: async (ctx) => {
    const inputs = ctx.optionValues as SetSensitivityLabelInputs;
    const core = getFxCore();
    const res = await core.setSensitivityLabel(inputs);
    return res;
  },
};

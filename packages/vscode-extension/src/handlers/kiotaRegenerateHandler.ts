// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, FxError, ok, Result, Stage, UserError } from "@microsoft/teamsfx-api";
import { ExtensionSource } from "../error/error";
import { localize } from "../utils/localizeUtils";
import { getSystemInputs } from "../utils/systemEnvUtils";
import { QuestionNames } from "@microsoft/teamsfx-core";
import { runCommand } from "./sharedOpts";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import { getTriggerFromProperty } from "../utils/telemetryUtils";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../telemetry/extTelemetryEvents";

export async function kiotaRegenerate(args?: any[]): Promise<Result<any, FxError>> {
  ExtTelemetry.sendTelemetryEvent(
    TelemetryEvent.KiotaRegenerateStart,
    getTriggerFromProperty(args)
  );
  if (!args || args.length !== 2) {
    const error = new UserError(
      ExtensionSource,
      "invalidParameter",
      localize("teamstoolkit.handler.createPluginWithManifest.error.missingParameter")
    );
    return err(error);
  }

  const specPath = args[0];
  const pluginManifestPath = args[1];
  const inputs = getSystemInputs();
  inputs[QuestionNames.ApiSpecLocation] = specPath;
  inputs[QuestionNames.ApiPluginManifestPath] = pluginManifestPath;

  const result = await runCommand(Stage.kiotaRegenerate, inputs);
  if (result.isErr()) {
    ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.KiotaRegenerate, result.error);
    return err(result.error);
  }

  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.KiotaRegenerate, {
    [TelemetryProperty.Success]: TelemetrySuccess.Yes,
  });
  return ok({});
}

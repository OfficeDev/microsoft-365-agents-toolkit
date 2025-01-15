// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  CreateProjectResult,
  err,
  FxError,
  ok,
  Result,
  Stage,
  UserError,
} from "@microsoft/teamsfx-api";
import { getSystemInputs } from "../utils/systemEnvUtils";
import {
  ApiPluginStartOptions,
  CapabilityOptions,
  KiotaLastCommands,
  ProjectTypeOptions,
  QuestionNames,
} from "@microsoft/teamsfx-core";
import { runCommand } from "./sharedOpts";
import * as vscode from "vscode";
import { openFolder } from "../utils/workspaceUtils";
import { ExtensionSource } from "../error/error";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import { getTriggerFromProperty } from "../utils/telemetryUtils";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../telemetry/extTelemetryEvents";
import { localize } from "../utils/localizeUtils";

export async function createPluginWithApiSpec(args?: any[]): Promise<Result<any, FxError>> {
  ExtTelemetry.sendTelemetryEvent(
    TelemetryEvent.CreatePluginWithManifestStart,
    getTriggerFromProperty(args)
  );
  if (!args || args.length > 99) {
    const error = new UserError(
      ExtensionSource,
      "invalidParameter",
      "Invalid parameter for creating plugin with API spec."
    );
    ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.CreatePluginWithManifest, error);
    return err(error);
  }

  const specPath = args[0];

  const inputs = getSystemInputs();
  inputs[QuestionNames.ApiSpecLocation] = specPath;
  inputs[QuestionNames.ApiPluginType] = ApiPluginStartOptions.apiSpec().id;
  inputs.capabilities = CapabilityOptions.declarativeAgent().id;
  inputs[QuestionNames.WithPlugin] = "yes";
  inputs[QuestionNames.ProjectType] = ProjectTypeOptions.Agent().id;

  const result = await runCommand(Stage.create, inputs);

  if (result.isErr()) {
    ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.CreatePluginWithManifest, result.error);
    return err(result.error);
  }

  const res = result.value as CreateProjectResult;
  const projectPathUri = vscode.Uri.file(res.projectPath);
  await openFolder(projectPathUri, true, res.warnings);
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.CreatePluginWithManifest, {
    [TelemetryProperty.Success]: TelemetrySuccess.Yes,
  });

  return ok({});
}

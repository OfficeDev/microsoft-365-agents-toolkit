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

export async function createDeclarativeAgentWithApiSpec(
  args?: any[]
): Promise<Result<any, FxError>> {
  ExtTelemetry.sendTelemetryEvent(
    TelemetryEvent.CreateDeclarativeAgentWithApiSpecStart,
    getTriggerFromProperty(args)
  );
  if (!args || args.length !== 1 || !args[0] || typeof args[0] !== "string") {
    const error = new UserError(
      ExtensionSource,
      "invalidParameter",
      localize("teamstoolkit.handler.createDeclarativeAgentWithApiSpec.error.invalidParameter")
    );
    ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.CreateDeclarativeAgentWithApiSpec, error);
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
    ExtTelemetry.sendTelemetryErrorEvent(
      TelemetryEvent.CreateDeclarativeAgentWithApiSpec,
      result.error
    );
    return err(result.error);
  }

  const res = result.value as CreateProjectResult;
  const projectPathUri = vscode.Uri.file(res.projectPath);
  await openFolder(projectPathUri, true, res.warnings);
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.CreateDeclarativeAgentWithApiSpec, {
    [TelemetryProperty.Success]: TelemetrySuccess.Yes,
  });

  return ok({});
}

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
import {
  ActionStartOptions,
  DACapabilityOptions,
  KiotaLastCommands,
  ProjectTypeOptions,
  QuestionNames,
} from "@microsoft/teamsfx-core";
import * as vscode from "vscode";
import { ExtensionSource } from "../error/error";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../telemetry/extTelemetryEvents";
import { localize } from "../utils/localizeUtils";
import { getSystemInputs } from "../utils/systemEnvUtils";
import { getTriggerFromProperty } from "../utils/telemetryUtils";
import { openFolder } from "../utils/workspaceUtils";
import { runCommand } from "./sharedOpts";

export async function createPluginWithManifest(args?: any[]): Promise<Result<any, FxError>> {
  ExtTelemetry.sendTelemetryEvent(
    TelemetryEvent.CreatePluginWithManifestStart,
    getTriggerFromProperty(args)
  );
  if (
    !args ||
    args.length < 3 ||
    args.length > 4 ||
    !args[2].lastCommand ||
    !Object.values(KiotaLastCommands).includes(args[2].lastCommand)
  ) {
    const error = new UserError(
      ExtensionSource,
      "invalidParameter",
      localize("teamstoolkit.handler.createPluginWithManifest.error.missingParameter")
    );
    ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.CreatePluginWithManifest, error);
    return err(error);
  }

  const specPath = args[0];
  const pluginManifestPath = args[1];
  const lastCommand = args[2].lastCommand;
  const outputFolder = args[3] ?? undefined;

  const inputs = getSystemInputs();
  inputs[QuestionNames.ApiSpecLocation] = specPath;
  inputs[QuestionNames.ActionManifestPath] = pluginManifestPath;
  inputs[QuestionNames.ActionType] = ActionStartOptions.apiSpec().id;
  inputs[QuestionNames.ApiOperation] = pluginManifestPath;
  let result;

  if (args[2].manifestPath && args[2].lastCommand === KiotaLastCommands.addPlugin) {
    // For add plugin
    inputs[QuestionNames.TeamsAppManifestFilePath] = args[2].manifestPath;
    result = await runCommand(Stage.addPlugin, inputs);

    if (result.isErr()) {
      ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.AddPluginWithManifest, result.error);
      return err(result.error);
    }

    ExtTelemetry.sendTelemetryEvent(TelemetryEvent.AddPluginWithManifest, {
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
    });
  } else {
    if (lastCommand === KiotaLastCommands.createDeclarativeCopilotWithManifest) {
      inputs.capabilities = DACapabilityOptions.declarativeAgent().id;
      inputs[QuestionNames.WithPlugin] = "yes";
    } else {
      inputs.capabilities = "api-plugin";
    }
    inputs[QuestionNames.ProjectType] = ProjectTypeOptions.copilotAgentOptionId;
    inputs[QuestionNames.Folder] = outputFolder;
    result = await runCommand(Stage.create, inputs);

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
  }
  return ok({});
}

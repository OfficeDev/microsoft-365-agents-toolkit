// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, Result, err, ok } from "@microsoft/teamsfx-api";
import { Hub, QuestionNames, assembleError } from "@microsoft/teamsfx-core";
import * as vscode from "vscode";
import { openHubWebClient } from "../debug/launch";
import { selectAndDebug } from "../debug/runIconHandler";
import { showError } from "../error/common";
import { core } from "../globalVariables";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../telemetry/extTelemetryEvents";
import { getSystemInputs } from "../utils/systemEnvUtils";
import { getTriggerFromProperty } from "../utils/telemetryUtils";
import { processResult } from "./sharedOpts";

export function debugInTestToolHandler(source: "treeview" | "message") {
  return async () => {
    if (source === "treeview") {
      ExtTelemetry.sendTelemetryEvent(TelemetryEvent.TreeViewDebugInTestTool);
    } else {
      ExtTelemetry.sendTelemetryEvent(TelemetryEvent.MessageDebugInTestTool);
    }
    await vscode.commands.executeCommand(
      "workbench.action.quickOpen",
      "debug Debug in Microsoft 365 Agents Playground"
    );
    return ok<unknown, FxError>(null);
  };
}

export async function selectAndDebugHandler(args?: any[]): Promise<Result<null, FxError>> {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.RunIconDebugStart, getTriggerFromProperty(args));
  const result = await selectAndDebug();
  await processResult(TelemetryEvent.RunIconDebug, result);
  return result;
}

export async function treeViewLocalDebugHandler(): Promise<Result<null, FxError>> {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.TreeViewLocalDebug);
  await vscode.commands.executeCommand("workbench.action.quickOpen", "debug ");
  return ok(null);
}

export async function treeViewPreviewHandler(...args: any[]): Promise<Result<null, FxError>> {
  ExtTelemetry.sendTelemetryEvent(
    TelemetryEvent.TreeViewPreviewStart,
    getTriggerFromProperty(args)
  );
  const properties: { [key: string]: string } = {};

  try {
    const env = args[1]?.identifier as string;
    const inputs = getSystemInputs();
    inputs.env = env;
    properties[TelemetryProperty.Env] = env;

    const result = await core.previewWithManifest(inputs);
    if (result.isErr()) {
      throw result.error;
    }

    const hub = inputs[QuestionNames.M365Host] as Hub;
    const url = result.value;
    properties[TelemetryProperty.Hub] = hub;

    await openHubWebClient(hub, url);
  } catch (error) {
    const assembledError = assembleError(error);
    void showError(assembledError);
    ExtTelemetry.sendTelemetryErrorEvent(
      TelemetryEvent.TreeViewPreview,
      assembledError,
      properties
    );
    return err(assembledError);
  }

  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.TreeViewPreview, {
    [TelemetryProperty.Success]: TelemetrySuccess.Yes,
    ...properties,
  });
  return ok(null);
}

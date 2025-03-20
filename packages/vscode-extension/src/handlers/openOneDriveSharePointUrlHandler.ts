// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as vscode from "vscode";
import { VS_CODE_UI } from "../qm/vsc_ui";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import {
  TelemetryTriggerFrom,
  TelemetryEvent,
  TelemetryProperty,
} from "../telemetry/extTelemetryEvents";

export async function openOneDriveSharePointUrlHandler(args: any[]): Promise<void> {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.OpenOneDriveSharePointUrlStart, {
    [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.Other,
  });
  const editor = vscode.window.activeTextEditor;
  if (!editor) {
    return;
  }

  await VS_CODE_UI.openUrl(args[0]);
}

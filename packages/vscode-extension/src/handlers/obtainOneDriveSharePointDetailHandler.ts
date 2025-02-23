// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as vscode from "vscode";
import { VS_CODE_UI } from "../qm/vsc_ui";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import {
  TelemetryTriggerFrom,
  TelemetrySuccess,
  TelemetryEvent,
  TelemetryProperty,
} from "../telemetry/extTelemetryEvents";
import { localize } from "../utils/localizeUtils";
import { getSystemInputs } from "../utils/systemEnvUtils";
import { core } from "../globalVariables";

export async function obtainOneDriveSharePointDetailHandler(args: any[]): Promise<void> {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.ObtainOneDriveSharePointDetailStart, {
    [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.Other,
  });
  const editor = vscode.window.activeTextEditor;
  if (!editor) {
    return;
  }

  const siteId = args[0];
  const itemId = args[1];
  const inputs = getSystemInputs();
  const result = await core.getODSPItemDetails(siteId, itemId, inputs);
  if (result.isOk()) {
    const res = await VS_CODE_UI.selectOption({
      title: "OneDrive Sharepoint Details",
      options: [result.value.name],
      name: "OneDrive Sharepoint Details",
    });
    if (res.isOk()) {
      void VS_CODE_UI.openUrl(result.value.url as string);
    } else {
      void vscode.window.showErrorMessage(res.error.message);
    }
  } else {
    ExtTelemetry.sendTelemetryErrorEvent(
      TelemetryEvent.ObtainOneDriveSharePointDetail,
      result.error
    );
    void vscode.window.showErrorMessage(result.error.message);
  }
}

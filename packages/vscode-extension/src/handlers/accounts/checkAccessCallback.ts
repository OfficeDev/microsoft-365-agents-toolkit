// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result, FxError, ok } from "@microsoft/teamsfx-api";
import { localize } from "../../utils/localizeUtils";
import { VS_CODE_UI } from "../../qm/vsc_ui";
import { ExtTelemetry } from "../../telemetry/extTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetryTriggerFrom,
} from "../../telemetry/extTelemetryEvents";
import { WebviewPanel } from "../../controls/webviewPanel";
import { PanelType } from "../../controls/PanelType";

export async function checkCopilotCallback(args?: any[]): Promise<Result<null, FxError>> {
  VS_CODE_UI.showMessage(
    "warn",
    localize("teamstoolkit.accountTree.copilotMessage"),
    false,
    localize("teamstoolkit.accountTree.copilotEnroll")
  )
    .then(async (result) => {
      if (result.isOk() && result.value === localize("teamstoolkit.accountTree.copilotEnroll")) {
        await VS_CODE_UI.openUrl("https://aka.ms/PluginsEarlyAccess");
        ExtTelemetry.sendTelemetryEvent(TelemetryEvent.OpenCopilotEnroll);
      }
    })
    .catch((_error) => {});
  return Promise.resolve(ok(null));
}

export function checkSideloadingCallback(args?: any[]): Promise<Result<null, FxError>> {
  VS_CODE_UI.showMessage(
    "error",
    localize("teamstoolkit.accountTree.sideloadingMessage"),
    false,
    localize("teamstoolkit.accountTree.sideloadingLearnMore")
  )
    .then((result) => {
      if (
        result.isOk() &&
        result.value === localize("teamstoolkit.accountTree.sideloadingLearnMore")
      ) {
        WebviewPanel.createOrShow(PanelType.AccountHelp);
        ExtTelemetry.sendTelemetryEvent(TelemetryEvent.OpenSideloadingLearnMore);
      }
    })
    .catch((_error) => {});
  WebviewPanel.createOrShow(PanelType.AccountHelp);
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.InteractWithInProductDoc, {
    [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.SideloadingDisabled,
  });
  return Promise.resolve(ok(null));
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, ok, Result } from "@microsoft/teamsfx-api";
import { PanelType } from "../../controls/PanelType";
import { WebviewPanel } from "../../controls/webviewPanel";
import { VS_CODE_UI } from "../../qm/vsc_ui";
import { ExtTelemetry } from "../../telemetry/extTelemetry";
import { TelemetryEvent } from "../../telemetry/extTelemetryEvents";
import { localize } from "../../utils/localizeUtils";

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
    localize("teamstoolkit.accountTree.sideloadingUseTestTenant"),
    localize("teamstoolkit.accountTree.sideloadingEnable")
  )
    .then(async (result) => {
      if (
        result.isOk() &&
        result.value === localize("teamstoolkit.accountTree.sideloadingEnable")
      ) {
        await VS_CODE_UI.openUrl(
          "https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/tools-prerequisites#enable-custom-app-upload-using-admin-center"
        );
        ExtTelemetry.sendTelemetryEvent(TelemetryEvent.OpenTestTenantLink);
      } else if (
        result.isOk() &&
        result.value === localize("teamstoolkit.accountTree.sideloadingUseTestTenant")
      ) {
        WebviewPanel.createOrShow(PanelType.AccountHelp);
        ExtTelemetry.sendTelemetryEvent(TelemetryEvent.OpenSideloadingEnable);
      }
    })
    .catch((_error) => {});
  return Promise.resolve(ok(null));
}

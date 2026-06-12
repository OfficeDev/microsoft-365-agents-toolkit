// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, ok, Result } from "@microsoft/teamsfx-api";
import { PanelType } from "../../controls/PanelType";
import { WebviewPanel } from "../../controls/webviewPanel";
import { VS_CODE_UI } from "../../qm/vsc_ui";
import { ExtTelemetry } from "../../telemetry/extTelemetry";
import { TelemetryEvent } from "../../telemetry/extTelemetryEvents";
import { localize } from "../../utils/localizeUtils";
import { commands } from "vscode";

export const checkAccessCallbackDeps = {
  showMessage: (
    messageLevel: "warn" | "error" | "info",
    message: string,
    modal: boolean,
    ...items: string[]
  ) => VS_CODE_UI.showMessage(messageLevel, message, modal, ...items),
  openUrl: (url: string) => VS_CODE_UI.openUrl(url),
  sendTelemetryEvent: (eventName: string) => ExtTelemetry.sendTelemetryEvent(eventName as any),
  createOrShow: (panelType: PanelType) => WebviewPanel.createOrShow(panelType),
  executeCommand: (command: string, ...args: any[]) => commands.executeCommand(command, ...args),
  localize: (key: string, ...args: any[]) => localize(key, ...args),
};

export async function checkCopilotCallback(args?: any[]): Promise<Result<null, FxError>> {
  checkAccessCallbackDeps
    .showMessage(
      "warn",
      checkAccessCallbackDeps.localize("teamstoolkit.accountTree.copilotMessage"),
      false,
      checkAccessCallbackDeps.localize("teamstoolkit.accountTree.copilotEnroll")
    )
    .then(async (result) => {
      if (
        result.isOk() &&
        result.value === checkAccessCallbackDeps.localize("teamstoolkit.accountTree.copilotEnroll")
      ) {
        await checkAccessCallbackDeps.openUrl(
          "https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/prerequisites"
        );
        checkAccessCallbackDeps.sendTelemetryEvent(TelemetryEvent.OpenCopilotEnroll);
      }
    })
    .catch((_error) => {});
  return Promise.resolve(ok(null));
}

export function checkSideloadingCallback(args?: any[]): Promise<Result<null, FxError>> {
  checkAccessCallbackDeps
    .showMessage(
      "error",
      checkAccessCallbackDeps.localize("teamstoolkit.accountTree.sideloadingMessage"),
      false,
      checkAccessCallbackDeps.localize("teamstoolkit.accountTree.sideloadingUseTestTenant"),
      checkAccessCallbackDeps.localize("teamstoolkit.accountTree.sideloadingEnable")
    )
    .then(async (result) => {
      if (
        result.isOk() &&
        result.value ===
          checkAccessCallbackDeps.localize("teamstoolkit.accountTree.sideloadingEnable")
      ) {
        await checkAccessCallbackDeps.openUrl(
          "https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/tools-prerequisites#enable-custom-app-upload-using-admin-center"
        );
        checkAccessCallbackDeps.sendTelemetryEvent(TelemetryEvent.OpenTestTenantLink);
      } else if (
        result.isOk() &&
        result.value ===
          checkAccessCallbackDeps.localize("teamstoolkit.accountTree.sideloadingUseTestTenant")
      ) {
        checkAccessCallbackDeps.createOrShow(PanelType.AccountHelp);
        checkAccessCallbackDeps.sendTelemetryEvent(TelemetryEvent.OpenSideloadingEnable);
      }
    })
    .catch((_error) => {});
  return Promise.resolve(ok(null));
}

/**
 * Suggest users to use sandboxed team for debugging
 * @param args
 * @returns
 */
export function checkSandboxCallback(args?: any[]): Promise<Result<null, FxError>> {
  checkAccessCallbackDeps
    .showMessage(
      "warn",
      checkAccessCallbackDeps.localize("teamstoolkit.accountTree.suggestSandboxedTeam"),
      false,
      checkAccessCallbackDeps.localize("teamstoolkit.accountTree.sandboxedTeam.button")
    )
    .then(async (result) => {
      if (
        result.isOk() &&
        result.value ===
          checkAccessCallbackDeps.localize("teamstoolkit.accountTree.sandboxedTeam.button")
      ) {
        await checkAccessCallbackDeps.executeCommand(
          "workbench.action.quickOpen",
          "debug Debug in sandbox in Teams (Edge)"
        );
      }
    })
    .catch((_error) => {});
  return Promise.resolve(ok(null));
}

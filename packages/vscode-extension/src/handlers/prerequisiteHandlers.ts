// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * @author Huajie Zhang <zhjay23@qq.com>
 */
"use strict";

import { FxError, Result, ok } from "@microsoft/teamsfx-api";
import { DepsManager, DepsType, assembleError } from "@microsoft/teamsfx-core";
import path from "path";
import * as vscode from "vscode";
import { checkPrerequisitesForGetStarted } from "../debug/depsChecker/getStartedChecker";
import { vscodeLogger } from "../debug/depsChecker/vscodeLogger";
import { vscodeTelemetry } from "../debug/depsChecker/vscodeTelemetry";
import { showError } from "../error/common";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import { TelemetryEvent } from "../telemetry/extTelemetryEvents";
import { acpInstalled } from "../utils/commonUtils";
import { localize } from "../utils/localizeUtils";
import { getTriggerFromProperty } from "../utils/telemetryUtils";

/**
 * Check required prerequisites in Get Started Page.
 */
export async function validateGetStartedPrerequisitesHandler(
  ...args: unknown[]
): Promise<Result<unknown, FxError>> {
  ExtTelemetry.sendTelemetryEvent(
    TelemetryEvent.ClickValidatePrerequisites,
    getTriggerFromProperty(args)
  );
  const result = await checkPrerequisitesForGetStarted();
  if (result.isErr()) {
    void showError(result.error);
    // // return non-zero value to let task "exit ${command:xxx}" to exit
    // return "1";
  }
  return result;
}

/**
 * Get path delimiter
 * Usage like ${workspaceFolder}/devTools/func${command:...}${env:PATH}
 */
export function getPathDelimiterHandler(): string {
  return path.delimiter;
}

/**
 * Get dotnet path to be referenced by task definition.
 * Usage like ${command:...}${env:PATH} so need to include delimiter as well
 */
export async function getDotnetPathHandler(): Promise<string> {
  try {
    const depsManager = new DepsManager(vscodeLogger, vscodeTelemetry);
    const dotnetStatus = (await depsManager.getStatus([DepsType.Dotnet]))?.[0];
    if (dotnetStatus?.isInstalled && dotnetStatus?.details?.binFolders !== undefined) {
      return `${path.delimiter}${dotnetStatus.details.binFolders
        .map((f: string) => path.dirname(f))
        .join(path.delimiter)}${path.delimiter}`;
    }
  } catch (error: any) {
    void showError(assembleError(error));
  }

  return `${path.delimiter}`;
}

export async function installAdaptiveCardExt(
  ...args: unknown[]
): Promise<Result<unknown, FxError>> {
  ExtTelemetry.sendTelemetryEvent(
    TelemetryEvent.AdaptiveCardPreviewerInstall,
    getTriggerFromProperty(args)
  );
  if (acpInstalled()) {
    await vscode.window.showInformationMessage(
      localize("teamstoolkit.handlers.adaptiveCardExtUsage")
    );
  } else {
    const selection = await vscode.window.showInformationMessage(
      localize("teamstoolkit.handlers.installAdaptiveCardExt"),
      "Install",
      "Cancel"
    );
    if (selection === "Install") {
      ExtTelemetry.sendTelemetryEvent(
        TelemetryEvent.AdaptiveCardPreviewerInstallConfirm,
        getTriggerFromProperty(args)
      );
      await vscode.commands.executeCommand(
        "workbench.extensions.installExtension",
        "TeamsDevApp.vscode-adaptive-cards"
      );
    } else {
      ExtTelemetry.sendTelemetryEvent(
        TelemetryEvent.AdaptiveCardPreviewerInstallCancel,
        getTriggerFromProperty(args)
      );
    }
  }
  return Promise.resolve(ok(null));
}

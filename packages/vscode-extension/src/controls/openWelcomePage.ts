// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { globalStateGet, globalStateUpdate } from "@microsoft/teamsfx-core";
import * as vscode from "vscode";
import { TelemetryTriggerFrom } from "../telemetry/extTelemetryEvents";
import { openWelcomeHandler } from "../handlers/controlHandlers";

const welcomePageKey = "ms-teams-vscode-extension.welcomePage.shown";

// Dependency injection wrapper for testability
export const openWelcomePageDeps = {
  globalStateGet: (key: string, defaultValue?: unknown) => globalStateGet(key, defaultValue),
  globalStateUpdate: (key: string, value: unknown) => globalStateUpdate(key, value),
};

export async function openWelcomePageAfterExtensionInstallation(): Promise<void> {
  if (await openWelcomePageDeps.globalStateGet(welcomePageKey, false)) {
    // Don't show: already showed
    return;
  }

  // Let's show!
  await openWelcomePageDeps.globalStateUpdate(welcomePageKey, true);
  await openWelcomeHandler(TelemetryTriggerFrom.Auto);
  await vscode.commands.executeCommand("workbench.view.extension.teamsfx");
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { globalStateGet, globalStateUpdate } from "@microsoft/teamsfx-core";
import * as vscode from "vscode";
import { TelemetryTriggerFrom } from "../telemetry/extTelemetryEvents";
import { openBuildIntelligentAppsWalkthroughHandler } from "../handlers/walkthrough";
import { openWelcomeHandler } from "../handlers/controlHandlers";

const welcomePageKey = "ms-teams-vscode-extension.welcomePage.shown";

export async function openWelcomePageAfterExtensionInstallation(): Promise<void> {
  if (await globalStateGet(welcomePageKey, false)) {
    // Don't show: already showed
    return;
  }

  // Let's show!
  await globalStateUpdate(welcomePageKey, true);
  await openWelcomeHandler([TelemetryTriggerFrom.Auto]);
  await openBuildIntelligentAppsWalkthroughHandler([TelemetryTriggerFrom.Auto]);
  await vscode.commands.executeCommand("workbench.view.extension.teamsfx");
}

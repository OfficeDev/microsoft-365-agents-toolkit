// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ok } from "@microsoft/teamsfx-api";
import { globalStateGet, globalStateUpdate } from "@microsoft/teamsfx-core";
import { GlobalKey, CommandKey } from "../constants";
import { workspaceUri } from "../globalVariables";
import { TelemetryTriggerFrom } from "../telemetry/extTelemetryEvents";
import {
  autoInstallDependencyHandler,
  showLocalDebugMessage,
  ShowScaffoldingWarningSummary,
} from "../utils/autoOpenHelper";
import { updateProjectStatus } from "../utils/projectStatusUtils";
import { openWelcomeHandler } from "./controlHandlers";
import { openReadMeHandler, openSampleReadmeHandler } from "./readmeHandlers";

export async function autoOpenProjectHandler(): Promise<void> {
  const isOpenWalkThrough = (await globalStateGet(GlobalKey.OpenWalkThrough, false)) as boolean;
  const isOpenReadMe = (await globalStateGet(GlobalKey.OpenReadMe, "")) as string;
  const isOpenSampleReadMe = (await globalStateGet(GlobalKey.OpenSampleReadMe, false)) as boolean;
  const createWarnings = (await globalStateGet(GlobalKey.CreateWarnings, "")) as string;
  const autoInstallDependency = (await globalStateGet(GlobalKey.AutoInstallDependency)) as boolean;
  if (isOpenWalkThrough) {
    await showLocalDebugMessage();
    await globalStateUpdate(GlobalKey.OpenWalkThrough, false);

    if (workspaceUri?.fsPath) {
      await ShowScaffoldingWarningSummary(workspaceUri.fsPath, createWarnings);
      await globalStateUpdate(GlobalKey.CreateWarnings, "");
    }
  }
  if (isOpenReadMe === workspaceUri?.fsPath) {
    await showLocalDebugMessage();
    await openReadMeHandler(TelemetryTriggerFrom.Auto);
    await updateProjectStatus(workspaceUri.fsPath, CommandKey.OpenReadMe, ok(null));
    await globalStateUpdate(GlobalKey.OpenReadMe, "");

    await ShowScaffoldingWarningSummary(workspaceUri.fsPath, createWarnings);
    await globalStateUpdate(GlobalKey.CreateWarnings, "");
  }
  if (isOpenSampleReadMe) {
    await showLocalDebugMessage();
    await openSampleReadmeHandler([TelemetryTriggerFrom.Auto]);
    await globalStateUpdate(GlobalKey.OpenSampleReadMe, false);
  }
  if (autoInstallDependency) {
    await autoInstallDependencyHandler();
    await globalStateUpdate(GlobalKey.AutoInstallDependency, false);
  }
}

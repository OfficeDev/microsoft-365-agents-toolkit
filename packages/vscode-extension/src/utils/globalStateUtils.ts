// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Warning } from "@microsoft/teamsfx-api";
import * as globalState from "@microsoft/teamsfx-core/build/common/globalState";
import { Uri } from "vscode";
import { GlobalKey } from "../constants";
import * as globalVariables from "../globalVariables";
import * as telemetryUtils from "./telemetryUtils";

export const globalStateUtilsDeps = {
  globalStateUpdate: globalState.globalStateUpdate,
  checkIsSPFx: globalVariables.checkIsSPFx,
  isTriggerFromWalkThrough: telemetryUtils.isTriggerFromWalkThrough,
};

export async function updateAutoOpenGlobalKey(
  showLocalDebugMessage: boolean,
  projectUri: Uri,
  warnings: Warning[] | undefined,
  args?: any[]
): Promise<void> {
  if (globalStateUtilsDeps.isTriggerFromWalkThrough(args)) {
    await globalStateUtilsDeps.globalStateUpdate(GlobalKey.OpenWalkThrough, true);
    await globalStateUtilsDeps.globalStateUpdate(GlobalKey.OpenReadMe, "");
  } else {
    await globalStateUtilsDeps.globalStateUpdate(GlobalKey.OpenWalkThrough, false);
    await globalStateUtilsDeps.globalStateUpdate(GlobalKey.OpenReadMe, projectUri.fsPath);
  }

  if (showLocalDebugMessage) {
    await globalStateUtilsDeps.globalStateUpdate(GlobalKey.ShowLocalDebugMessage, true);
  }

  if (warnings?.length) {
    await globalStateUtilsDeps.globalStateUpdate(
      GlobalKey.CreateWarnings,
      JSON.stringify(warnings)
    );
  }

  if (globalStateUtilsDeps.checkIsSPFx(projectUri.fsPath)) {
    void globalStateUtilsDeps.globalStateUpdate(GlobalKey.AutoInstallDependency, true);
  }
}

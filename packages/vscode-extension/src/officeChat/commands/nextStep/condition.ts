// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CommandKey } from "../../../constants";
import { RecordedActions } from "../../../utils/projectStatusUtils";
import { OfficeWholeStatus } from "./types";

/**
 * if some Teams App is opened in the workspace
 * @param status
 * @returns
 */
export function isProjectOpened(status: OfficeWholeStatus): boolean {
  return !!status.projectOpened;
}

export function isNodeInstalled(status: OfficeWholeStatus): boolean {
  return !!status.projectOpened && !!status.projectOpened.isNodeInstalled;
}

/**
 * if did no action after the project is scaffolded
 * @param status
 * @returns
 */
export function isDidNoActionAfterScaffolded(status: OfficeWholeStatus): boolean {
  const actionStatus = status.projectOpened?.actionStatus;
  if (actionStatus) {
    for (const key of RecordedActions) {
      if (actionStatus[key].result !== "no run") {
        return false;
      }
    }
  }

  return true;
}

/**
 * if the source code is modified after the last debug succeeded
 * @param status
 * @returns
 */
export function isDebugSucceededAfterSourceCodeChanged(status: OfficeWholeStatus): boolean {
  if (!status.projectOpened) {
    return false;
  }
  return (
    status.projectOpened.actionStatus[CommandKey.LocalDebug].result === "success" &&
    status.projectOpened.actionStatus[CommandKey.LocalDebug].time >
      status.projectOpened.codeModifiedTime.source
  );
}

/**
 * if the Office Add-in can be previewed in the local environment
 * @param status
 * @returns
 */
export function canOfficeAddInPreviewInLocalEnv(status: OfficeWholeStatus): boolean {
  return (
    !!status.projectOpened &&
    !!status.projectOpened.launchJSONContent &&
    (status.projectOpened.launchJSONContent.toLocaleLowerCase().includes("desktop (edge legacy)") ||
      status.projectOpened.launchJSONContent
        .toLocaleLowerCase()
        .includes("desktop (edge chromium)"))
  );
}

/**
 * if node_modules exists to check whether dependencies are installed
 * @param status
 * @returns
 */
export function isDependenciesInstalled(status: OfficeWholeStatus): boolean {
  return !!status.projectOpened && !!status.projectOpened.nodeModulesExist;
}

/**
 * if there is a readme file in the project
 * @param status
 * @returns
 */
export function isHaveReadMe(status: OfficeWholeStatus): boolean {
  return !!status.projectOpened && !!status.projectOpened.readmeContent;
}

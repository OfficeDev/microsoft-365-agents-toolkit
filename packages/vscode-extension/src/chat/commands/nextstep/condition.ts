// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CommandKey } from "../../../constants";
import { RecordedActions } from "../../../utils/projectStatusUtils";
import { WholeStatus } from "./types";

/**
 * if Teams Toolkit is first installed
 * @param status
 * @returns
 */
export function isFirstInstalled(status: WholeStatus): boolean {
  return status.machineStatus.firstInstalled;
}

/**
 * if some Teams App is opened in the workspace
 * @param status
 * @returns
 */
export function isProjectOpened(status: WholeStatus): boolean {
  return !!status.projectOpened;
}

/**
 * if did no action after the project is scaffolded
 * @param status
 * @returns
 */
export function isDidNoActionAfterScaffolded(status: WholeStatus): boolean {
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
export function isDebugSucceededAfterSourceCodeChanged(status: WholeStatus): boolean {
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
 * if can preview in the test tool
 * @param status
 * @returns
 */
export function canPreviewInTestTool(status: WholeStatus): boolean {
  return (
    !!status.projectOpened &&
    !!status.projectOpened.launchJSONContent &&
    status.projectOpened.launchJSONContent.toLocaleLowerCase().includes("test tool")
  );
}

/**
 * if user has logged in M365 account
 * @param status
 * @returns
 */
export function isM365AccountLogin(status: WholeStatus): boolean {
  return status.machineStatus.m365LoggedIn;
}

/**
 * if provision is succeeded after the infra code is changed
 * @param status
 * @returns
 */
export function isProvisionedSucceededAfterInfraCodeChanged(status: WholeStatus): boolean {
  return (
    !!status.projectOpened &&
    status.projectOpened.actionStatus[CommandKey.Provision].result === "success" &&
    status.projectOpened.actionStatus[CommandKey.Provision].time >
      status.projectOpened.codeModifiedTime.infra
  );
}

/**
 * if user has logged in Azure account
 * @param status
 * @returns
 */
export function isAzureAccountLogin(status: WholeStatus): boolean {
  return status.machineStatus.azureLoggedIn;
}

/**
 * if deploy is succeeded after the source code is changed
 * @param status
 * @returns
 */
export function isDeployedAfterSourceCodeChanged(status: WholeStatus): boolean {
  return (
    !!status.projectOpened &&
    status.projectOpened.actionStatus[CommandKey.Deploy].result === "success" &&
    status.projectOpened.actionStatus[CommandKey.Deploy].time >
      status.projectOpened.codeModifiedTime.source
  );
}

/**
 * if publish is succeeded once
 * @param status
 * @returns
 */
export function isPublishedSucceededBefore(status: WholeStatus): boolean {
  return (
    !!status.projectOpened &&
    status.projectOpened.actionStatus[CommandKey.Publish].result === "success"
  );
}

/**
 * if there is a readme file in the project
 * @param status
 * @returns
 */
export function isHaveReadMe(status: WholeStatus): boolean {
  return !!status.projectOpened && !!status.projectOpened.readmeContent;
}

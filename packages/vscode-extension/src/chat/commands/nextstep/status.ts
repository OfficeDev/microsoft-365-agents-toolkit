// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as projectStatusUtils from "../../../utils/projectStatusUtils";
import * as helper from "./helper";
import { MachineStatus, WholeStatus } from "./types";

export const firstInstalledKey = "first-installation";

export const statusDeps = {
  getProjectMetadata: helper.getProjectMetadata,
  getProjectStatus: projectStatusUtils.getProjectStatus,
  getFileModifiedTime: projectStatusUtils.getFileModifiedTime,
  getREADME: projectStatusUtils.getREADME,
  getLaunchJSON: projectStatusUtils.getLaunchJSON,
  globalStateGet: helper.globalStateGet,
  globalStateUpdate: helper.globalStateUpdate,
  checkCredential: helper.checkCredential,
};

export async function getWholeStatus(folder?: string): Promise<WholeStatus> {
  if (!folder) {
    return {
      machineStatus: await getMachineStatus(),
    };
  } else {
    const projectSettings = statusDeps.getProjectMetadata(folder);
    const projectId = projectSettings?.projectId;
    const actionStatus = await statusDeps.getProjectStatus(projectId ?? folder);
    const codeModifiedTime = {
      source: await statusDeps.getFileModifiedTime(
        `${folder.split("\\").join("/")}/**/*.{ts,tsx,js,jsx}`
      ),
      infra: await statusDeps.getFileModifiedTime(`${folder.split("\\").join("/")}/infra/**/*`),
    };

    return {
      machineStatus: await getMachineStatus(),
      projectOpened: {
        path: folder,
        projectId,
        codeModifiedTime,
        readmeContent: await statusDeps.getREADME(folder),
        actionStatus,
        launchJSONContent: await statusDeps.getLaunchJSON(folder),
      },
    };
  }
}

export async function getMachineStatus(): Promise<MachineStatus> {
  const firstInstalled = await statusDeps.globalStateGet(firstInstalledKey, true);
  await statusDeps.globalStateUpdate(firstInstalledKey, false);
  return {
    firstInstalled,
    ...(await statusDeps.checkCredential()),
  };
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { OfficeWholeStatus } from "./types";
import * as Status from "../../../chat/commands/nextstep/status";
import fs from "fs-extra";
import child_process from "child_process";

export async function getWholeStatus(folder?: string): Promise<OfficeWholeStatus> {
  return Status.getWholeStatus(folder).then(async (status: OfficeWholeStatus) => {
    if (status.projectOpened) {
      if (folder !== undefined) {
        status.projectOpened.nodeModulesExist = await fs.pathExists(`${folder}/node_modules`);
        status.projectOpened.isNodeInstalled = await checkNodeInstallation();
      }
    }
    return status;
  });
}

function checkNodeInstallation(): Promise<boolean> {
  return new Promise((resolve) => {
    child_process.exec("node -v", (error) => {
      if (error) {
        resolve(false);
        return;
      }
      resolve(true);
    });
  });
}

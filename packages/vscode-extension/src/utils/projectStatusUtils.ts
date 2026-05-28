// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ConfigFolderName, Result } from "@microsoft/teamsfx-api";
import { FeatureFlags, featureFlagManager } from "@microsoft/teamsfx-core";
import fs from "fs-extra";
import { glob } from "glob";
import * as os from "os";
import { getProjectMetadata } from "../chat/commands/nextstep/helper";
import { ProjectActionStatus } from "../chat/commands/nextstep/types";
import { CommandKey } from "../constants";

export const projectStatusFilePath = os.homedir() + `/.${ConfigFolderName}/projectStates.json`;

export const RecordedActions: (keyof ProjectActionStatus)[] = [
  CommandKey.Provision,
  CommandKey.Deploy,
  CommandKey.Publish,
  CommandKey.OpenReadMe,
];

export function emptyProjectStatus(): ProjectActionStatus {
  return {
    [CommandKey.LocalDebug]: { result: "no run", time: new Date(0) },
    [CommandKey.Provision]: { result: "no run", time: new Date(0) },
    [CommandKey.Deploy]: { result: "no run", time: new Date(0) },
    [CommandKey.Publish]: { result: "no run", time: new Date(0) },
    [CommandKey.OpenReadMe]: { result: "no run", time: new Date(0) },
  };
}

export async function getProjectStatus(projectId: string): Promise<ProjectActionStatus> {
  let status = emptyProjectStatus();
  if (await fs.pathExists(projectStatusFilePath)) {
    try {
      const content = await fs.readFile(projectStatusFilePath, "utf8");
      const json = JSON.parse(content, (_, value) => {
        const date = Date.parse(value);
        if (!isNaN(date)) {
          return new Date(date);
        } else {
          return value;
        }
      });
      status = { ...status, ...json[projectId] };
    } catch {}
  }
  return status;
}

export async function updateProjectStatus(
  fsPath: string,
  commandName: string,
  result: Result<unknown, Error>,
  forced = false
) {
  const projectSettings = getProjectMetadata(fsPath);
  const p = projectSettings?.projectId ?? fsPath;
  const actions = RecordedActions.map((x) => x.toString());
  if (actions.includes(commandName) || forced) {
    /// save project action running status
    const status = await getProjectStatus(p);
    status[commandName as keyof ProjectActionStatus] = {
      result: result.isOk() ? "success" : "fail",
      time: new Date(Date.now()),
    };
    let json: any = {};
    if (await fs.pathExists(projectStatusFilePath)) {
      try {
        json = JSON.parse(await fs.readFile(projectStatusFilePath, "utf8"));
      } catch {}
    }
    try {
      json[p] = status;
      await fs.writeFile(projectStatusFilePath, JSON.stringify(json, null, 2));
    } catch {}
  }
}

export async function getFileModifiedTime(pattern: string): Promise<Date> {
  const files = await glob(pattern, { ignore: "node_modules/**" });
  let lastModifiedTime = new Date(0);
  for (const file of files) {
    const stat = await fs.stat(file);
    if (stat.mtime > lastModifiedTime) {
      lastModifiedTime = stat.mtime;
    }
  }
  return lastModifiedTime;
}

export async function getREADME(folder: string): Promise<string | undefined> {
  const readmePath = `${folder}/README.md`;
  if (await fs.pathExists(readmePath)) {
    return await fs.readFile(readmePath, "utf-8");
  }
  return undefined;
}

export async function getLaunchJSON(folder: string): Promise<string | undefined> {
  const launchJSONPath = `${folder}/.vscode/launch.json`;
  if (await fs.pathExists(launchJSONPath)) {
    return await fs.readFile(launchJSONPath, "utf-8");
  }
  return undefined;
}

export function getWalkThroughId(): string {
  return featureFlagManager.getBooleanValue(FeatureFlags.ChatParticipantUIEntries)
    ? "TeamsDevApp.ms-teams-vscode-extension#teamsToolkitGetStartedWithChat"
    : "TeamsDevApp.ms-teams-vscode-extension#teamsToolkitGetStarted";
}

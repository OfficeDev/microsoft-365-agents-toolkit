// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as vscode from "vscode";
import fs from "fs-extra";
import { isValidProject } from "@microsoft/teamsfx-core";
import { initializeGlobalVariables, context } from "../globalVariables";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import { TelemetryEvent, TelemetryProperty } from "../telemetry/extTelemetryEvents";
import TreeViewManagerInstance from "../treeview/treeViewManager";

export const fileSystemWatcherOps = {
  isValidProject: (workspacePath: string) => isValidProject(workspacePath),
  createFileSystemWatcher: (pattern: string) => vscode.workspace.createFileSystemWatcher(pattern),
  initializeGlobalVariables: () => initializeGlobalVariables(context),
  updateDevelopmentTreeView: () => TreeViewManagerInstance.updateDevelopmentTreeView(),
  readJson: (filePath: string) => fs.readJson(filePath),
  sendTelemetryEvent: (eventName: string, properties?: any) =>
    ExtTelemetry.sendTelemetryEvent(eventName as any, properties),
};
const fileSystemWatcherDeps = fileSystemWatcherOps;

export function addFileSystemWatcher(workspacePath: string) {
  if (fileSystemWatcherDeps.isValidProject(workspacePath)) {
    const packageLockFileWatcher =
      fileSystemWatcherDeps.createFileSystemWatcher("**/package-lock.json");

    packageLockFileWatcher.onDidCreate(async (event) => {
      await sendSDKVersionTelemetry(event.fsPath);
    });

    packageLockFileWatcher.onDidChange(async (event) => {
      await sendSDKVersionTelemetry(event.fsPath);
    });

    const yorcFileWatcher = fileSystemWatcherDeps.createFileSystemWatcher("**/.yo-rc.json");
    yorcFileWatcher.onDidCreate((_event) => {
      refreshSPFxTreeOnFileChanged();
    });
    yorcFileWatcher.onDidChange((_event) => {
      refreshSPFxTreeOnFileChanged();
    });
    yorcFileWatcher.onDidDelete((_event) => {
      refreshSPFxTreeOnFileChanged();
    });
  }
}

export function refreshSPFxTreeOnFileChanged() {
  fileSystemWatcherDeps.initializeGlobalVariables();
  fileSystemWatcherDeps.updateDevelopmentTreeView();
}

export async function sendSDKVersionTelemetry(filePath: string) {
  const packageLockFile = (await fileSystemWatcherDeps.readJson(filePath).catch(() => {})) as {
    dependencies: { [key: string]: { version: string } };
  };
  fileSystemWatcherDeps.sendTelemetryEvent(TelemetryEvent.UpdateSDKPackages, {
    [TelemetryProperty.BotbuilderVersion]: packageLockFile?.dependencies["botbuilder"]?.version,
    [TelemetryProperty.TeamsFxVersion]:
      packageLockFile?.dependencies["@microsoft/teamsfx"]?.version,
    [TelemetryProperty.TeamsJSVersion]:
      packageLockFile?.dependencies["@microsoft/teams-js"]?.version,
  });
}

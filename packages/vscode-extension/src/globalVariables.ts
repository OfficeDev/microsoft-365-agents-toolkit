// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs from "fs-extra";
import path from "path";
import * as vscode from "vscode";

import { UserState } from "./constants";
import {
  FxCore,
  isValidProject,
  isValidOfficeAddInProject,
  isManifestOnlyOfficeAddinProject,
  manifestUtils,
} from "@microsoft/teamsfx-core";
import { TeamsAppManifest, Tools } from "@microsoft/teamsfx-api";

/**
 * Common variables used throughout the extension. They must be initialized in the activate() method of extension.ts
 */
export let context: vscode.ExtensionContext;
export let workspaceUri: vscode.Uri | undefined;
export let isTeamsFxProject = false;
export let isOfficeAddInProject = false;
export let isOfficeManifestOnlyProject = false;
export let isSPFxProject = false;
export let isDeclarativeCopilotApp = false;
export let isExistingUser = "no";
export let defaultExtensionLogPath: string;
export let commandIsRunning = false;
export let core: FxCore;
export let tools: Tools;
export let diagnosticCollection: vscode.DiagnosticCollection; // Collection of diagnositcs after running app validation.
export let deleteAadInProgress = false;

export interface ILocalDebugPorts {
  checkPorts: number[];
  conflictPorts: number[];
  terminateButton: string;
  process2conflictPorts: Record<string, number[]>;
  terminateProcesses: string[];
}

export const LocalDebugPorts: ILocalDebugPorts = {
  checkPorts: [],
  conflictPorts: [],
  terminateButton: "",
  process2conflictPorts: {},
  terminateProcesses: [],
};

export function resetLocalDebugPorts() {
  LocalDebugPorts.checkPorts = [];
  LocalDebugPorts.conflictPorts = [];
  LocalDebugPorts.terminateButton = "";
  LocalDebugPorts.process2conflictPorts = {};
  LocalDebugPorts.terminateProcesses = [];
}

if (vscode.workspace && vscode.workspace.workspaceFolders) {
  if (vscode.workspace.workspaceFolders.length > 0) {
    workspaceUri = vscode.workspace.workspaceFolders[0].uri;
  }
}

export function initializeGlobalVariables(ctx: vscode.ExtensionContext): void {
  context = ctx;
  isExistingUser = context.globalState.get<string>(UserState.IsExisting) || "no";
  isTeamsFxProject = isValidProject(workspaceUri?.fsPath);
  isOfficeAddInProject = isValidOfficeAddInProject(workspaceUri?.fsPath);
  if (isOfficeAddInProject) {
    isOfficeManifestOnlyProject = isManifestOnlyOfficeAddinProject(workspaceUri?.fsPath);
  }
  // Default Extension log path
  // eslint-disable-next-line no-secrets/no-secrets
  // e.g. C:/Users/xx/AppData/Roaming/Code/logs/20230221T095340/window7/exthost/TeamsDevApp.ms-teams-vscode-extension
  defaultExtensionLogPath = ctx.logUri.fsPath;
  if (!fs.pathExistsSync(defaultExtensionLogPath)) {
    fs.mkdirSync(defaultExtensionLogPath);
  }
  if (isTeamsFxProject && workspaceUri?.fsPath) {
    isSPFxProject = checkIsSPFx(workspaceUri?.fsPath);
    isDeclarativeCopilotApp = checkIsDeclarativeCopilotApp(workspaceUri.fsPath);
  } else {
    isSPFxProject = fs.existsSync(path.join(workspaceUri?.fsPath ?? "./", "SPFx"));
  }
}

export function checkIsSPFx(directory: string): boolean {
  const files = fs.readdirSync(directory);
  for (const file of files) {
    if (file === ".yo-rc.json") {
      const content = fs.readJsonSync(path.join(directory, file)) as Record<string, unknown>;
      if (content["@microsoft/generator-sharepoint"]) {
        return true;
      }
    } else if (fs.lstatSync(path.join(directory, file)).isDirectory()) {
      if (checkIsSPFx(path.join(directory, file))) return true;
    }
  }
  return false;
}

export function checkIsDeclarativeCopilotApp(directory: string): boolean {
  const manifestRes = manifestUtils.readAppManifestSync(directory);
  if (manifestRes.isOk()) {
    return manifestUtils.getCapabilities(manifestRes.value).includes("copilotGpt");
  } else {
    return false;
  }
}

export function updateIsDeclarativeCopilotApp(manifest: TeamsAppManifest): boolean {
  const value = manifestUtils.getCapabilities(manifest).includes("copilotGpt");
  isDeclarativeCopilotApp = value;
  return isDeclarativeCopilotApp;
}

export function setCommandIsRunning(isRunning: boolean) {
  commandIsRunning = isRunning;
}

// Only used by checkProjectUpgradable() when error happens
export function unsetIsTeamsFxProject() {
  isTeamsFxProject = false;
}

export function setTools(toolsInstance: Tools) {
  tools = toolsInstance;
}
export function setCore(coreInstance: FxCore) {
  core = coreInstance;
}

export function setDiagnosticCollection(collection: vscode.DiagnosticCollection) {
  diagnosticCollection = collection;
}

export function setDeleteAadInProgress(inProgress: boolean) {
  deleteAadInProgress = inProgress;
}

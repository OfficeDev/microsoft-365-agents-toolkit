// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { isValidProject } from "@microsoft/teamsfx-core";
import { workspaceUri, core } from "../globalVariables";
import { TelemetryProperty, TelemetryTriggerFrom } from "../telemetry/extTelemetryEvents";
import { getSystemInputs } from "./systemEnvUtils";

export function getPackageVersion(versionStr: string): string {
  if (versionStr.includes("alpha")) {
    return "alpha";
  }

  if (versionStr.includes("beta")) {
    return "beta";
  }

  if (versionStr.includes("rc")) {
    return "rc";
  }

  return "formal";
}

export async function getProjectId(): Promise<string | undefined> {
  if (!workspaceUri) {
    return undefined;
  }
  try {
    const ws = workspaceUri.fsPath;
    const projInfoRes = await core.getProjectId(ws);
    if (projInfoRes.isOk()) {
      return projInfoRes.value;
    }
  } catch (e) {}
  return undefined;
}

export function getTriggerFromProperty(args?: any[]) {
  // if not args are not supplied, by default, it is trigger from "CommandPalette"
  // e.g. vscode.commands.executeCommand("fx-extension.openWelcome");
  // in this case, "fx-exentiosn.openWelcome" is trigged from "CommandPalette".
  if (!args || args.length === 0 || !args[0]) {
    return { [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.CommandPalette };
  }

  switch ((args[0] as TelemetryTriggerFrom).toString()) {
    case TelemetryTriggerFrom.TreeView:
      return { [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.TreeView };
    case TelemetryTriggerFrom.ViewTitleNavigation:
      return { [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.ViewTitleNavigation };
    case TelemetryTriggerFrom.QuickPick:
      return { [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.QuickPick };
    case TelemetryTriggerFrom.Webview:
      return { [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.Webview };
    case TelemetryTriggerFrom.CodeLens:
      return { [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.CodeLens };
    case TelemetryTriggerFrom.EditorTitle:
      return { [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.EditorTitle };
    case TelemetryTriggerFrom.SideBar:
      return { [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.SideBar };
    case TelemetryTriggerFrom.Notification:
      return { [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.Notification };
    case TelemetryTriggerFrom.WalkThrough:
      return { [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.WalkThrough };
    case TelemetryTriggerFrom.CopilotChat:
      return { [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.CopilotChat };
    case TelemetryTriggerFrom.Auto:
      return { [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.Auto };
    case TelemetryTriggerFrom.ExternalUrl:
      return { [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.ExternalUrl };
    case TelemetryTriggerFrom.Other:
      return { [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.Other };
    case TelemetryTriggerFrom.CreateAppQuestionFlow:
      return { [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.CreateAppQuestionFlow };
    default:
      return { [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.Unknow };
  }
}

export function isTriggerFromWalkThrough(args?: any[]): boolean {
  if (!args || args.length === 0) {
    return false;
  } else if (
    (args[0] as TelemetryTriggerFrom).toString() === TelemetryTriggerFrom.WalkThrough ||
    (args[0] as TelemetryTriggerFrom).toString() === TelemetryTriggerFrom.Notification
  ) {
    return true;
  }

  return false;
}

export interface TeamsAppTelemetryInfo {
  appId: string;
  tenantId: string;
}

export async function getTeamsAppTelemetryInfoByEnv(
  env: string
): Promise<TeamsAppTelemetryInfo | undefined> {
  try {
    const ws = workspaceUri!.fsPath;
    if (isValidProject(ws)) {
      const projectInfoRes = await core.getProjectInfo(ws, env);
      if (projectInfoRes.isOk()) {
        const projectInfo = projectInfoRes.value;
        return {
          appId: projectInfo.teamsAppId,
          tenantId: projectInfo.m365TenantId,
        };
      }
    }
  } catch (e) {}
  return undefined;
}

export async function getSettingsVersion(): Promise<string | undefined> {
  if (core) {
    const versionCheckResult = await projectVersionCheck();

    if (versionCheckResult.isOk()) {
      return versionCheckResult.value.currentVersion;
    }
  }
  return undefined;
}

export async function projectVersionCheck() {
  return await core.projectVersionCheck(getSystemInputs());
}

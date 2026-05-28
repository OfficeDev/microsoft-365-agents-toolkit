// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import fs from "fs-extra";
import path from "path";
import { MetadataV3, telemetryUtils } from "@microsoft/teamsfx-core";
import { core, workspaceUri } from "../globalVariables";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import { ConfigFolderName } from "@microsoft/teamsfx-api";

export async function checkProjectTypeAndSendTelemetry(): Promise<void> {
  if (!workspaceUri?.fsPath) return;
  const res = await core.checkProjectType(workspaceUri?.fsPath);
  if (res.isErr()) return;
  const result = res.value;
  const props: Record<string, string> = {};
  telemetryUtils.fillinProjectTypeProperties(props, result);
  for (const key of Object.keys(props)) {
    ExtTelemetry.addSharedProperty(key, props[key]);
  }
}

// Only work in ts/js project
export function isTestToolEnabledProject(workspacePath: string): boolean {
  const testToolYmlPath = path.join(workspacePath, MetadataV3.testToolConfigFile);
  if (fs.pathExistsSync(testToolYmlPath)) {
    return true;
  }
  return false;
}

export async function isM365Project(workspacePath: string): Promise<boolean> {
  const projectSettingsPath = path.resolve(
    workspacePath,
    `.${ConfigFolderName}`,
    "configs",
    "projectSettings.json"
  );

  if (await fs.pathExists(projectSettingsPath)) {
    const projectSettings = await fs.readJson(projectSettingsPath);
    return projectSettings.isM365;
  } else {
    return false;
  }
}

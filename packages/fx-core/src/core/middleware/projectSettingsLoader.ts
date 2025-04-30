// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs-extra";
import * as path from "path";
import * as uuid from "uuid";

import { FxError, Inputs, Result, Stage, StaticPlatforms, ok } from "@microsoft/teamsfx-api";

import { isVSProject } from "../../common/projectSettingsHelper";
import {
  Component,
  TelemetryEvent,
  TelemetryProperty,
  sendTelemetryEvent,
} from "../../common/telemetry";
import { MetadataV2, MetadataV3 } from "../../common/versionMetadata";
import { convertProjectSettingsV2ToV3 } from "../../component/migrate";
import { globalVars } from "../../common/globalVars";
import { CoreHookContext } from "../types";
import { pathUtils } from "../../component/utils/pathUtils";

// export this for V2 -> V3 migration purpose
export async function loadProjectSettingsByProjectPathV2(
  projectPath: string
): Promise<Result<any, FxError>> {
  const settingsFile = getProjectSettingPathV2(projectPath);
  const projectSettings: any = await fs.readJson(settingsFile);
  if (!projectSettings.projectId) {
    projectSettings.projectId = uuid.v4();
    sendTelemetryEvent(Component.core, TelemetryEvent.FillProjectId, {
      [TelemetryProperty.ProjectId]: projectSettings.projectId,
    });
  }
  globalVars.isVS = isVSProject(projectSettings);
  return ok(convertProjectSettingsV2ToV3(projectSettings, projectPath));
}

export function shouldIgnored(ctx: CoreHookContext): boolean {
  const inputs = ctx.arguments[ctx.arguments.length - 1] as Inputs;
  const method = ctx.method;

  let isCreate = false;
  if (method === "getQuestions") {
    const task = ctx.arguments[0] as Stage;
    isCreate = task === Stage.create;
  }

  return StaticPlatforms.includes(inputs.platform) || isCreate || inputs.ignoreLockByUT;
}

export function getProjectSettingsPath(projectPath: string): string {
  return pathUtils.getYmlFilePath(projectPath, "dev", true) || "";
}

export function getProjectSettingPathV2(projectPath: string): string {
  return path.resolve(projectPath, MetadataV2.configFolder, "configs", MetadataV2.configFile);
}

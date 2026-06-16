// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import path from "path";
import semver from "semver";
import {
  MetadataV2,
  MetadataV3,
  MetadataV3Abandoned,
  VersionInfo,
  VersionSource,
  VersionState,
} from "../../../common/versionMetadata";
import { settingsUtil } from "../../../component/utils/settingsUtil";
import { CoreHookContext } from "../../types";
import { getProjectSettingPathV2, getProjectSettingsPath } from "../projectSettingsLoader";

export async function getProjectVersion(ctx: CoreHookContext): Promise<VersionInfo> {
  const projectPath = getParameterFromCxt(ctx, "projectPath", "");
  return await getProjectVersionFromPath(projectPath);
}

export async function getProjectVersionFromPath(projectPath: string): Promise<VersionInfo> {
  const v3path = getProjectSettingsPath(projectPath);
  if (v3path && (await fs.pathExists(v3path))) {
    const readSettingsResult = await settingsUtil.readSettings(projectPath, false);
    if (readSettingsResult.isOk()) {
      return {
        version: readSettingsResult.value.version || "",
        source: VersionSource.teamsapp,
      };
    }
    throw readSettingsResult.error;
  }

  const v2path = getProjectSettingPathV2(projectPath);
  if (await fs.pathExists(v2path)) {
    const settings = await fs.readJson(v2path);
    return {
      version: settings.version || "",
      source: VersionSource.projectSettings,
    };
  }

  const abandonedPath = path.resolve(
    projectPath,
    MetadataV3Abandoned.configFolder,
    MetadataV3Abandoned.configFile
  );
  if (await fs.pathExists(abandonedPath)) {
    return {
      version: MetadataV3Abandoned.configFolder,
      source: VersionSource.settings,
    };
  }

  return {
    version: "",
    source: VersionSource.unknown,
  };
}

export async function getTrackingIdFromPath(projectPath: string): Promise<string> {
  const v3path = getProjectSettingsPath(projectPath);
  if (await fs.pathExists(v3path)) {
    const readSettingsResult = await settingsUtil.readSettings(projectPath, false);
    return readSettingsResult.isOk() ? readSettingsResult.value.trackingId : "";
  }

  const v2path = getProjectSettingPathV2(projectPath);
  if (await fs.pathExists(v2path)) {
    const settings = await fs.readJson(v2path);
    if (settings.projectId) {
      return settings.projectId || "";
    }
  }

  return "";
}

export function getVersionState(info: VersionInfo): VersionState {
  if (
    info.source === VersionSource.projectSettings &&
    semver.gte(info.version, MetadataV2.projectVersion) &&
    semver.lte(info.version, MetadataV2.projectMaxVersion)
  ) {
    return VersionState.upgradeable;
  }
  if (
    info.source === VersionSource.teamsapp &&
    semver.valid(info.version) &&
    semver.lt(info.version, MetadataV3.unSupprotVersion)
  ) {
    return VersionState.compatible;
  }
  if (info.source === VersionSource.teamsapp && !semver.valid(info.version)) {
    return VersionState.compatible;
  }
  return VersionState.unsupported;
}

export function getParameterFromCxt(
  ctx: CoreHookContext,
  key: string,
  defaultValue?: string
): string {
  const inputs = ctx.arguments[ctx.arguments.length - 1] as Inputs;
  const value = (inputs[key] as string) || defaultValue || "";
  return value;
}

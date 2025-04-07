// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Context,
  err,
  FxError,
  InputsWithProjectPath,
  Result,
  TeamsAppManifest,
  UserError,
  ok,
} from "@microsoft/teamsfx-api";
import axios from "axios";
import fs from "fs-extra";
import { DriverContext } from "../component/driver/interface/commonArgs";
import { getLocalizedString } from "./localizeUtils";
import { pathUtils } from "../component/utils/pathUtils";
import { metadataUtil } from "../component/utils/metadataUtil";
import { resolve } from "../component/configManager/lifecycle";
import { DriverDefinition } from "../component/configManager/interface";
import * as path from "path";
import AdmZip from "adm-zip";
import { Constants } from "../component/driver/teamsApp/constants";

export async function waitSeconds(second: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, second * 1000));
}

export function generateDriverContext(ctx: Context, inputs: InputsWithProjectPath): DriverContext {
  return {
    azureAccountProvider: ctx.tokenProvider!.azureAccountProvider,
    m365TokenProvider: ctx.tokenProvider!.m365TokenProvider,
    ui: ctx.userInteraction,
    progressBar: undefined,
    logProvider: ctx.logProvider,
    telemetryReporter: ctx.telemetryReporter,
    projectPath: inputs.projectPath,
    platform: inputs.platform,
  };
}

export async function isJsonSpecFile(filePath: string): Promise<boolean> {
  const specPath = filePath.toLowerCase();
  if (specPath.endsWith(".yaml") || specPath.endsWith(".yml")) {
    return false;
  } else if (specPath.endsWith(".json")) {
    return true;
  }
  const isRemoteFile = specPath.startsWith("http:") || specPath.startsWith("https:");

  try {
    const fileContent = isRemoteFile
      ? (await axios.get(specPath)).data
      : await fs.readFile(specPath, "utf-8");
    JSON.parse(fileContent);
    return true;
  } catch (error) {
    return false;
  }
}

// Read teamsapp.yaml and get the value of teamsapp id, shared title id, and shared app id
// Output [teamsapp id, shared title id, shared app id]
export async function parseShareAppActionYamlConfig(
  projectPath: string
): Promise<Result<string[], FxError>> {
  const templatePath = pathUtils.getYmlFilePath(projectPath, "dev");
  const maybeProjectModel = await metadataUtil.parse(templatePath, "dev");
  if (maybeProjectModel.isErr()) {
    return err(maybeProjectModel.error);
  }
  const projectModel = maybeProjectModel.value;
  if (!projectModel.share || !projectModel.share.driverDefs) {
    return err(
      new UserError(
        "FxCore",
        "Share to Users",
        getLocalizedString("error.share.yamlConfigNotFound")
      )
    );
  }
  const shareToOthersAction = projectModel.share.driverDefs.find(
    (d) => d.uses === "teamsApp/shareToOthers"
  );
  if (!shareToOthersAction) {
    return err(
      new UserError(
        "FxCore",
        "Share to Users",
        getLocalizedString("error.share.shareActionConfigNotFound", "teamsApp/shareToOthers")
      )
    );
  }
  // 1. get manifest id
  const appPackagePath = (shareToOthersAction.with as any)?.appPackagePath;
  if (!appPackagePath) {
    return err(
      new UserError(
        "FxCore",
        "Share to Users",
        getLocalizedString("error.share.appPackageConfigNotFound")
      )
    );
  }
  const resolvedDriver = resolve(shareToOthersAction, [], []) as DriverDefinition;
  const resolvedAppPackagePath = path.resolve(
    projectPath,
    (resolvedDriver.with as any).appPackagePath as string
  );
  const zipEntries = new AdmZip(resolvedAppPackagePath).getEntries();
  const manifestFile = zipEntries.find((x) => x.entryName === Constants.MANIFEST_FILE);
  if (!manifestFile) {
    return err(
      new UserError(
        "FxCore",
        "Share to Users",
        getLocalizedString("error.share.manifestFileNotFound")
      )
    );
  }
  const manifest = JSON.parse(manifestFile.getData().toString()) as TeamsAppManifest;
  const manifestId = manifest.id;
  if (!manifestId) {
    return err(
      new UserError(
        "FxCore",
        "Share to Users",
        getLocalizedString("error.share.manifestIdNotFound")
      )
    );
  }

  // 2. get shared title id and shared app id
  const sharedTitleIdEnvName = (shareToOthersAction.writeToEnvironmentFile as any)?.titleId;
  const sharedAppIdEnvName = (shareToOthersAction.writeToEnvironmentFile as any)?.appId;
  if (!sharedTitleIdEnvName || !sharedAppIdEnvName) {
    return err(
      new UserError(
        "FxCore",
        "Share to Users",
        getLocalizedString("error.share.sharedConfigNotFound")
      )
    );
  }
  // env file has already been loaded before calling this function.
  const sharedTitleId = process.env[sharedTitleIdEnvName];
  const sharedAppId = process.env[sharedAppIdEnvName];
  if (!sharedTitleId || !sharedAppId) {
    return err(
      new UserError(
        "FxCore",
        "Share to Users",
        getLocalizedString("error.share.sharedIdNotFound", sharedTitleId, sharedAppId)
      )
    );
  }
  return ok([manifestId, sharedTitleId, sharedAppId]);
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, FxError, ok, Result, TeamsAppManifest, UserError } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import fs from "fs-extra";
import path from "path";
import { getLocalizedString } from "../../../common/localizeUtils";
import { DriverDefinition } from "../../configManager/interface";
import { resolve } from "../../configManager/lifecycle";
import { envUtil } from "../../utils/envUtil";
import { metadataUtil } from "../../utils/metadataUtil";
import { pathUtils } from "../../utils/pathUtils";
import { Constants } from "../teamsApp/constants";
import * as shareToOthers from "./shareToOthers";

// Read m365agents.yaml and get the value of shared title id, and shared app id
export async function parseShareAppActionYamlConfig(
  projectPath: string
): Promise<Result<{ teamsappId: string; titleId: string; appId: string }, FxError>> {
  const templatePath = pathUtils.getYmlFilePath(projectPath, "dev") as string;
  const maybeProjectModel = await metadataUtil.parse(templatePath);
  if (maybeProjectModel.isErr()) {
    return err(maybeProjectModel.error);
  }
  const projectModel = maybeProjectModel.value;
  if (!projectModel.deploy || !projectModel.deploy.driverDefs) {
    return err(
      new UserError("FxCore", "Share", getLocalizedString("error.share.yamlConfigNotFound"))
    );
  }
  const shareToOthersAction = projectModel.deploy.driverDefs.find(
    (d) => d.uses === shareToOthers.actionName
  );
  if (!shareToOthersAction) {
    return err(
      new UserError(
        "FxCore",
        "Share",
        getLocalizedString("error.share.shareActionConfigNotFound", shareToOthers.actionName)
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

  const readEnvRes = await envUtil.readEnv(projectPath, "dev");
  if (readEnvRes.isErr()) {
    return err(readEnvRes.error);
  }
  const resolvedDriver = resolve(shareToOthersAction, [], []) as DriverDefinition;
  const resolvedAppPackagePath = path.resolve(
    projectPath,
    (resolvedDriver.with as any).appPackagePath as string
  );
  if (!fs.existsSync(resolvedAppPackagePath)) {
    return err(
      new UserError(
        "FxCore",
        "Share",
        getLocalizedString("error.share.appPackageNotFound", resolvedAppPackagePath)
      )
    );
  }
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
      new UserError("FxCore", "Share", getLocalizedString("error.share.sharedConfigNotFound"))
    );
  }
  // env file has already been loaded before calling this function.
  const sharedTitleId = process.env[sharedTitleIdEnvName];
  const sharedAppId = process.env[sharedAppIdEnvName];
  if (!sharedTitleId || !sharedAppId) {
    return err(
      new UserError(
        "FxCore",
        "Share",
        getLocalizedString("error.share.sharedIdNotFound", sharedTitleId, sharedAppId)
      )
    );
  }
  return ok({ teamsappId: manifestId, titleId: sharedTitleId, appId: sharedAppId });
}

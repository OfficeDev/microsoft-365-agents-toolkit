// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { AppManifestUtils, FxError, Result, TeamsManifest, ok } from "@microsoft/teamsfx-api";
import "reflect-metadata";
import { TOOLS } from "../common/globalVars";

export async function generateConfigFiles(
  appManifestFilePath: string,
  programmingLanguage: string,
  includePlayground: boolean,
  includeLocalDebug: boolean,
  includeRemoteDeploy: boolean
): Promise<Result<undefined, FxError>> {
  const appManifest = await AppManifestUtils.readTeamsManifest(appManifestFilePath);
  const appName = appManifest.name.short;

  if (includePlayground) {
    if (isPlaygroundSupported(appManifest)) {
      await TOOLS.ui?.showMessage(
        "info",
        `Generating Playground configuration files for ${appName} in ${programmingLanguage}`,
        false
      );
    } else {
      await TOOLS.ui?.showMessage(
        "warn",
        `Playground is not supported for the current app manifest of ${appName}. Skipping Playground configuration file generation.`,
        false
      );
    }
  }

  return ok(undefined);
}

function isPlaygroundSupported(manifest: TeamsManifest): boolean {
  if (manifest.bots && manifest.bots.length > 0) {
    return true;
  }
  return false;
}

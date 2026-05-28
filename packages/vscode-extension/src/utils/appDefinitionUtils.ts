// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { envUtil, metadataUtil, pathUtils } from "@microsoft/teamsfx-core";
import { core, workspaceUri } from "../globalVariables";
import { UserError } from "@microsoft/teamsfx-api";
import { ExtensionErrors, ExtensionSource } from "../error/error";

export async function getAppName(): Promise<string | undefined> {
  if (!workspaceUri) {
    return undefined;
  }
  try {
    const ws = workspaceUri.fsPath;
    const nameRes = await core.getTeamsAppName(ws);
    if (nameRes.isOk() && nameRes.value != "") {
      return nameRes.value;
    }
  } catch (e) {}
  return undefined;
}

export async function getV3TeamsAppId(projectPath: string, env: string): Promise<string> {
  const result = await envUtil.readEnv(projectPath, env, false);
  if (result.isErr()) {
    throw result.error;
  }

  const teamsAppIdKey = (await getTeamsAppKeyName(env)) || "TEAMS_APP_ID";
  const teamsAppId = result.value[teamsAppIdKey];
  if (teamsAppId === undefined) {
    throw new UserError(
      ExtensionSource,
      ExtensionErrors.TeamsAppIdNotFoundError,
      `TEAMS_APP_ID is missing in ${env} environment.`
    );
  }

  return teamsAppId;
}

export async function getTeamsAppKeyName(env?: string): Promise<string | undefined> {
  const templatePath = pathUtils.getYmlFilePath(workspaceUri!.fsPath, env);
  const maybeProjectModel = await metadataUtil.parse(templatePath, env);
  if (maybeProjectModel.isErr()) {
    return undefined;
  }
  const projectModel = maybeProjectModel.value;
  if (projectModel.provision?.driverDefs && projectModel.provision.driverDefs.length > 0) {
    for (const driver of projectModel.provision.driverDefs) {
      if (driver.uses === "teamsApp/create") {
        return driver.writeToEnvironmentFile?.teamsAppId;
      }
    }
  }
  return undefined;
}

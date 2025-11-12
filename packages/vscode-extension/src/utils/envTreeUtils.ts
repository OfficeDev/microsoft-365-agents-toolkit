// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { SubscriptionInfo } from "@microsoft/teamsfx-api";
import { dotenvUtil } from "@microsoft/teamsfx-core/build/src/component/utils/envUtil";
import fs from "fs-extra";
import path from "path";
import { workspaceUri } from "../globalVariables";
import { getV3TeamsAppId } from "./appDefinitionUtils";
import { getProvisionResultJson } from "./fileSystemUtils";

export async function getSubscriptionInfoFromEnv(
  env: string
): Promise<SubscriptionInfo | undefined> {
  let provisionResult: Record<string, any> | undefined;

  try {
    provisionResult = await getProvisionResultJson(env);
  } catch (error) {
    // ignore error on tree view when load provision result failed.
    return undefined;
  }

  if (!provisionResult) {
    return undefined;
  }

  if (provisionResult.solution && provisionResult.solution.subscriptionId) {
    return {
      subscriptionName: provisionResult.solution.subscriptionName,
      subscriptionId: provisionResult.solution.subscriptionId,
      tenantId: provisionResult.solution.tenantId,
    };
  } else {
    return undefined;
  }
}

export async function getM365TenantFromEnv(env: string): Promise<string | undefined> {
  const projectPath = workspaceUri!.fsPath;
  const envFile = path.resolve(projectPath, "env", `.env.${env}`);
  if (await fs.pathExists(envFile)) {
    const envData = dotenvUtil.deserialize(fs.readFileSync(envFile, "utf-8"));
    return envData.obj["TEAMS_APP_TENANT_ID"];
  }
  return undefined;
}

export async function getResourceGroupNameFromEnv(env: string): Promise<string | undefined> {
  let provisionResult: Record<string, any> | undefined;

  try {
    provisionResult = await getProvisionResultJson(env);
  } catch (error) {
    // ignore error on tree view when load provision result failed.
    return undefined;
  }

  if (!provisionResult) {
    return undefined;
  }

  return provisionResult.solution?.resourceGroupName;
}

export async function getProvisionSucceedFromEnv(env: string): Promise<boolean | undefined> {
  // If TEAMS_APP_ID is set, it's highly possible that the project is provisioned.
  try {
    const teamsAppId = await getV3TeamsAppId(workspaceUri!.fsPath, env);
    return teamsAppId !== "";
  } catch (error) {
    return false;
  }
}

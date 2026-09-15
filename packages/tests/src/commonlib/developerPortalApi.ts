// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { AxiosInstance } from "axios";

const newDeveloperPortalApisFlag = "TEAMSFX_NEW_DP_APIS";

interface AppListItem {
  appId: string;
  appExternalId?: string;
}

interface AppListResponse {
  items?: AppListItem[];
  continuationToken?: string;
}

export function isUsingNewDeveloperPortalApis(): boolean {
  const value = process.env[newDeveloperPortalApisFlag];
  return value === "1" || value?.toLowerCase() === "true";
}

export function getBotRegistrationApiPath(botId: string): string {
  return isUsingNewDeveloperPortalApis()
    ? `/v1.0/botregistrations/${botId}`
    : `/api/botframework/${botId}`;
}

export async function getTeamsAppApiPath(
  requester: AxiosInstance,
  teamsAppId: string,
): Promise<string> {
  if (!isUsingNewDeveloperPortalApis()) {
    return `/api/appdefinitions/${teamsAppId}`;
  }

  let continuationToken: string | undefined;
  do {
    const response = await requester.get<AppListResponse>("/v1.0/apps", {
      params: { pageSize: 100 },
      headers: continuationToken
        ? { "x-ms-continuation": continuationToken }
        : undefined,
    });
    const app = response.data.items?.find(
      (item) => item.appId === teamsAppId || item.appExternalId === teamsAppId,
    );
    if (app) {
      return `/v1.0/apps/${app.appId}`;
    }
    continuationToken = response.data.continuationToken;
  } while (continuationToken);

  throw new Error(`Cannot resolve Teams app resource ID for ${teamsAppId}`);
}

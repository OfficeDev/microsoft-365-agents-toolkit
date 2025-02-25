// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { Tunnel } from "@microsoft/dev-tunnels-contracts";
import {
  ManagementApiVersions,
  TunnelManagementHttpClient,
} from "@microsoft/dev-tunnels-management";
import { FxError, M365TokenProvider, Result, SystemError, err, ok } from "@microsoft/teamsfx-api";
import axios from "axios";
import { teamsDevPortalClient } from "../client/teamsDevPortalClient";
import { GraphReadUserScopes, SPFxScopes } from "./constants";
import fs from "fs-extra";
import path from "path";
import { MetadataV3 } from "./versionMetadata";

export async function getSideloadingStatus(token: string): Promise<boolean | undefined> {
  return teamsDevPortalClient.getSideloadingStatus(token);
}

export async function listAllTenants(token: string): Promise<Record<string, any>[]> {
  const RM_ENDPOINT = "https://management.azure.com/tenants?api-version=2022-06-01";
  if (token.length > 0) {
    try {
      const response = await axios.get(RM_ENDPOINT, {
        headers: { Authorization: `Bearer ${token}` },
      });
      return response.data.value;
    } catch (error) {
      return [];
    }
  }

  return [];
}

export async function getSPFxTenant(graphToken: string): Promise<string> {
  const GRAPH_TENANT_ENDPT = "https://graph.microsoft.com/v1.0/sites/root?$select=webUrl";
  if (graphToken.length > 0) {
    const response = await axios.get(GRAPH_TENANT_ENDPT, {
      headers: { Authorization: `Bearer ${graphToken}` },
    });
    return response.data.webUrl;
  }
  return "";
}

export async function getSPFxToken(
  m365TokenProvider: M365TokenProvider
): Promise<string | undefined> {
  const graphTokenRes = await m365TokenProvider.getAccessToken({
    scopes: GraphReadUserScopes,
  });
  let spoToken = undefined;
  if (graphTokenRes && graphTokenRes.isOk()) {
    const tenant = await getSPFxTenant(graphTokenRes.value);
    const spfxTokenRes = await m365TokenProvider.getAccessToken({
      scopes: SPFxScopes(tenant),
    });
    spoToken = spfxTokenRes.isOk() ? spfxTokenRes.value : undefined;
  }
  return spoToken;
}

// this function will be deleted after VS has added get dev tunnel and list dev tunnels API
const TunnelManagementUserAgent = { name: "Teams-Toolkit" };
export async function listDevTunnels(
  token: string,
  isGitHub = false
): Promise<Result<Tunnel[], FxError>> {
  try {
    const tunnelManagementClientImpl = new TunnelManagementHttpClient(
      TunnelManagementUserAgent,
      ManagementApiVersions.Version20230927preview,
      () => {
        if (isGitHub === true) {
          const res = `github client_id=a200baed193bb2088a6e ${token}`;
          return Promise.resolve(res);
        } else {
          const res = `Bearer ${token}`;
          return Promise.resolve(res);
        }
      }
    );

    const options = {
      includeAccessControl: true,
    };
    const tunnels = await tunnelManagementClientImpl.listTunnels(undefined, undefined, options);
    return ok(tunnels);
  } catch (error) {
    return err(new SystemError("DevTunnels", "ListDevTunnelsFailed", error.message));
  }
}

export function isTestToolEnabledProject(projectPath: string): boolean {
  const testToolYmlPath = path.join(projectPath, MetadataV3.testToolConfigFile);
  if (fs.pathExistsSync(testToolYmlPath)) {
    return true;
  }
  return false;
}

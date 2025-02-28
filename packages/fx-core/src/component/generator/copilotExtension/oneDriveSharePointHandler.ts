// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Context,
  FxError,
  Inputs,
  Result,
  SystemError,
  UserError,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import { GraphScopes } from "../../../common/constants";
import axios, { AxiosInstance } from "axios";
import { OneDriveSharePointItemType } from "../constant";

export interface ItemMetadata {
  id: string;
  name: string;
  uniqueId?: string;
  listId?: string;
  webId?: string;
  siteId?: string;
  webUrl?: string;
  itemType?: OneDriveSharePointItemType;
}

/**
 * Create a graph client with token
 * @param context The context
 * @returns The graph client
 */
export async function createGraphClientWithToken(context: Context): Promise<Result<AxiosInstance, FxError>> {
  const graphTokenRes = await context.tokenProvider?.m365TokenProvider.getAccessToken({
    scopes: GraphScopes,
  });
  if (!graphTokenRes?.isOk()) {
    return err(
      new SystemError({
        source: "copilotPlugin",
        name: "GetGraphTokenFailed",
        message: "Failed to get Graph token",
        displayMessage: "Failed to get Graph token",
      })
    );
  }
  const client = axios.create({
    baseURL: "https://graph.microsoft.com/v1.0",
    headers: { Authorization: `Bearer ${graphTokenRes.value}` },
  });
  return ok(client);
}

/**
 * Get the SharePoint site by relative path
 * @param graphClient The graph client
 * @param url The share point url
 * @returns The SharePoint site
 */
export async function getSharePointSiteByRelativePath(graphClient: AxiosInstance, url: string): Promise<Result<ItemMetadata, FxError>> {
  // Extract the hostname and relative path from the url
  const urlObj = new URL(url);
  const hostname = urlObj.hostname;
  const relativePath = urlObj.pathname;
  try {
    const res = await graphClient.get(`/sites/${hostname}:${relativePath}?$select=id,name,sharepointIds`);
    return ok({
      id: res.data.id,
      name: res.data.name,
      webId: res.data.sharepointIds.webId,
      siteId: res.data.sharepointIds.siteId,
    });
  } catch (error) {
    return err(new UserError({
      source: "copilotPlugin",
      name: "GetSharePointSiteFailed",
      message: "Failed to get SharePoint site",
      displayMessage: "Failed to get SharePoint site",
    }));
  }
}

/**
 * Encode the share point url
 * @param itemUrl The share point url
 * @returns The encoded url
 */
export function encodeSharePointUrl(itemUrl: string): string {
  const base64Value = Buffer.from(itemUrl).toString("base64");
  return "u!" + base64Value.replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
}

/**
 * Get the drive item info
 * @param graphClient The graph client
 * @param encodedUrl The encoded url
 * @returns The drive item info
 */
export async function getDriveItemInfo(
  graphClient: any,
  encodedUrl: string
): Promise<ItemMetadata> {
  const res = await graphClient.get(`/shares/${encodedUrl}/driveItem?$select=id,name,sharepointIds,webUrl,file,folder`);
  return {
    id: res.data.id,
    name: res.data.name,
    uniqueId: res.data.sharepointIds.listItemUniqueId,
    listId: res.data.sharepointIds.listId,
    webId: res.data.sharepointIds.webId,
    siteId: res.data.sharepointIds.siteId,
    webUrl: res.data.webUrl,
    itemType: res.data.file ? OneDriveSharePointItemType.File : OneDriveSharePointItemType.Folder,
  };
}

export async function getODSPItemDetailById(
  context: Context,
  siteId: string,
  itemId: string,
): Promise<Result<ItemMetadata[], UserError>> {
  const graphClientResult = await createGraphClientWithToken(context);
  if (graphClientResult.isErr()) {
    return err(graphClientResult.error);
  }
  const graphClient = graphClientResult.value;

  try {
    const itemRes = await graphClient.get(`/sites/${siteId}/drive/items/${itemId}`);
    return ok([
      {
        id: itemRes.data.id,
        name: itemRes.data.name,
        url: itemRes.data.webUrl,
      },
    ]);
  } catch (error) {
    if (axios.isAxiosError(error) && error.response) {
      if (error.response!.status >= 400 && error.response!.status < 510) {
        return err(new UserError(
          "getODSPItemDetailById",
          "GraphApiError",
          error.response.data.error.message,
          error.response.data.error.message
        ));
      }
    }
    return err(new SystemError(
      "getODSPItemDetailById",
      "GraphApiError",
      error.message,
      error.message
    ));
  }
}
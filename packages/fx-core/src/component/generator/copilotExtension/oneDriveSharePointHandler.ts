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
import axios from "axios";
import { OneDriveSharePointItemType } from "../constant";

export interface ItemMetadata {
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
export async function createGraphClientWithToken(context: Context): Promise<Result<any, FxError>> {
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
  inputs: Inputs
): Promise<Result<ItemMetadata[], UserError>> {
  const graphClientResult = await createGraphClientWithToken(context);
  if (graphClientResult.isErr()) {
    return err(graphClientResult.error);
  }
  const graphClient = graphClientResult.value;

  const itemRes = await graphClient.get(`/sites/${siteId}/drive/items/${itemId}`);

  return ok([
    {
      id: itemRes.data.id,
      label: itemRes.data.name,
      name: itemRes.data.name,
      url: itemRes.data.webUrl,
    },
  ]);
}
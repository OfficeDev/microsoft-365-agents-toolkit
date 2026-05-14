// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result, ok, err } from "neverthrow";
import { DeclarativeAgentManifestWrapper } from "@microsoft/app-manifest";
import type { AtkError } from "../../core/error";
import { systemError } from "../../core/error";
import type { ODSPItemMetadata } from "../types";

/** Parameters referencing an item by SharePoint IDs. */
export interface ODSPSharePointIds {
  site_id: string;
  web_id: string;
  list_id?: string;
  unique_id?: string;
}

/**
 * Add OneDrive / SharePoint knowledge capability to a DA manifest.
 *
 * When `item` is provided, it scopes the knowledge to that specific item.
 * When omitted, the capability enables access to all OneDrive/SharePoint content.
 */
export async function addOneDriveSharePointKnowledge(
  agentManifestPath: string,
  item?: ODSPItemMetadata
): Promise<Result<void, AtkError>> {
  try {
    const wrapper = await DeclarativeAgentManifestWrapper.read(agentManifestPath);

    if (item) {
      if (item.webUrl) {
        wrapper.addOneDriveSharePointCapability({
          items_by_url: [{ url: item.webUrl }],
        });
      } else if (item.siteId && item.webId) {
        wrapper.addOneDriveSharePointCapability({
          items_by_sharepoint_ids: [
            {
              site_id: item.siteId,
              web_id: item.webId,
              ...(item.listId && { list_id: item.listId }),
              ...(item.uniqueId && { unique_id: item.uniqueId }),
            },
          ],
        });
      }
    } else {
      wrapper.addOneDriveSharePointCapability();
    }

    await wrapper.save(agentManifestPath);
    return ok(undefined);
  } catch (e) {
    return err(
      systemError(
        "AddOneDriveSharePointKnowledgeFailed",
        `Failed to add OneDrive/SharePoint knowledge: ${e}`,
        {
          source: "declarativeAgent/knowledge",
          inner: e instanceof Error ? e : new Error(String(e)),
        }
      )
    );
  }
}

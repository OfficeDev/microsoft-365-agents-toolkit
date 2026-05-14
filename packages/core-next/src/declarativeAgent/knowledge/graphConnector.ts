// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result, ok, err } from "neverthrow";
import { DeclarativeAgentManifestWrapper } from "@microsoft/app-manifest";
import type { AtkError } from "../../core/error";
import { systemError } from "../../core/error";

/**
 * Add Graph Connector knowledge capability to a DA manifest.
 *
 * @param agentManifestPath  Absolute path to the declarative agent manifest.
 * @param connectionIds      One or more Graph Connector connection IDs to add.
 */
export async function addGraphConnectorKnowledge(
  agentManifestPath: string,
  connectionIds: string[]
): Promise<Result<void, AtkError>> {
  try {
    const wrapper = await DeclarativeAgentManifestWrapper.read(agentManifestPath);
    wrapper.addGraphConnectorsCapability(connectionIds);
    await wrapper.save(agentManifestPath);
    return ok(undefined);
  } catch (e) {
    return err(
      systemError(
        "AddGraphConnectorKnowledgeFailed",
        `Failed to add Graph Connector knowledge: ${e}`,
        {
          source: "declarativeAgent/knowledge",
          inner: e instanceof Error ? e : new Error(String(e)),
        }
      )
    );
  }
}

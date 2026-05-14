// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result, ok, err } from "neverthrow";
import { DeclarativeAgentManifestWrapper } from "@microsoft/app-manifest";
import type { AtkError } from "../../core/error";
import { systemError } from "../../core/error";

/**
 * Add web-search capability to a declarative agent manifest.
 *
 * When `siteUrl` is provided, the capability is scoped to that site.
 * When omitted, the capability enables general web search (all web).
 */
export async function addWebSearchKnowledge(
  agentManifestPath: string,
  siteUrl?: string
): Promise<Result<void, AtkError>> {
  try {
    const wrapper = await DeclarativeAgentManifestWrapper.read(agentManifestPath);

    if (siteUrl) {
      wrapper.addWebSearchCapability([{ url: siteUrl }]);
    } else {
      wrapper.addWebSearchCapability();
    }

    await wrapper.save(agentManifestPath);
    return ok(undefined);
  } catch (e) {
    return err(
      systemError("AddWebSearchKnowledgeFailed", `Failed to add web search knowledge: ${e}`, {
        source: "declarativeAgent/knowledge",
        inner: e instanceof Error ? e : new Error(String(e)),
      })
    );
  }
}

/**
 * Validate that the web-search capability would not produce a conflicting state.
 *
 * When switching from scoped→all-web, existing scoped sites will be removed.
 * Returns `true` if caller should prompt user for confirmation.
 */
export function webSearchRequiresConfirmation(
  agentManifestPath: string,
  newSiteUrl: string | undefined
): boolean {
  try {
    const wrapper = DeclarativeAgentManifestWrapper.readSync(agentManifestPath);
    if (newSiteUrl === undefined) {
      // Switching to all-web: check if scoped sites exist
      const caps = wrapper.capabilities;
      const webSearch = caps.find(
        (c) => "name" in c && (c as Record<string, unknown>)["name"] === "WebSearch"
      );
      if (webSearch && "sites" in webSearch) {
        const sites = (webSearch as Record<string, unknown>)["sites"];
        return Array.isArray(sites) && sites.length > 0;
      }
    }
    return false;
  } catch {
    return false;
  }
}

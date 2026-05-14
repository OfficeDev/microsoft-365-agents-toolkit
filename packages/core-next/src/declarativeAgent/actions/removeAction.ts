// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result, ok, err } from "neverthrow";
import { DeclarativeAgentManifestWrapper } from "@microsoft/app-manifest";
import type { AtkError } from "../../core/error";
import { userError, systemError } from "../../core/error";

/**
 * Remove an action from a declarative agent manifest by its action ID.
 *
 * @param agentManifestPath Absolute path to the declarative agent manifest.
 * @param actionId          The action ID to remove.
 */
export async function removeAction(
  agentManifestPath: string,
  actionId: string
): Promise<Result<void, AtkError>> {
  try {
    const wrapper = await DeclarativeAgentManifestWrapper.read(agentManifestPath);

    if (!wrapper.hasAction(actionId)) {
      return err(
        userError("ActionNotFound", `Action "${actionId}" not found in the agent manifest.`, {
          source: "declarativeAgent/actions",
        })
      );
    }

    wrapper.removeAction(actionId);
    await wrapper.save(agentManifestPath);
    return ok(undefined);
  } catch (e) {
    return err(
      systemError("RemoveActionFailed", `Failed to remove action: ${e}`, {
        source: "declarativeAgent/actions",
        inner: e instanceof Error ? e : new Error(String(e)),
      })
    );
  }
}

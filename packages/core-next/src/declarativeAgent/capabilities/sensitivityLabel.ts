// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result, ok, err } from "neverthrow";
import { DeclarativeAgentManifestWrapper } from "@microsoft/app-manifest";
import type { AtkError } from "../../core/error";
import { systemError } from "../../core/error";
import type { SetSensitivityLabelInput } from "../types";

/**
 * Set a sensitivity label on a declarative agent manifest.
 *
 * Reads the manifest, sets `sensitivity_label.id`, and writes it back.
 */
export async function setSensitivityLabel(
  input: SetSensitivityLabelInput
): Promise<Result<void, AtkError>> {
  try {
    const wrapper = await DeclarativeAgentManifestWrapper.read(input.agentManifestPath);
    const jsonStr = wrapper.toJSON();
    const manifest = JSON.parse(jsonStr) as Record<string, unknown>;

    manifest.sensitivity_label = { id: input.labelId };

    const updated = DeclarativeAgentManifestWrapper.fromJSON(JSON.stringify(manifest, null, 2));
    await updated.save(input.agentManifestPath);
    return ok(undefined);
  } catch (e) {
    return err(
      systemError("SetSensitivityLabelFailed", `Failed to set sensitivity label: ${e}`, {
        source: "declarativeAgent/capabilities",
        inner: e instanceof Error ? e : new Error(String(e)),
      })
    );
  }
}

/**
 * Add or update conversation starters in a declarative agent manifest.
 *
 * @param agentManifestPath Absolute path to the agent manifest.
 * @param starters          Array of conversation starter objects with `text` and optional `title`.
 */
export async function setConversationStarters(
  agentManifestPath: string,
  starters: Array<{ text: string; title?: string }>
): Promise<Result<void, AtkError>> {
  try {
    const wrapper = await DeclarativeAgentManifestWrapper.read(agentManifestPath);
    for (const s of starters) {
      wrapper.addConversationStarter(s.text, s.title);
    }
    await wrapper.save(agentManifestPath);
    return ok(undefined);
  } catch (e) {
    return err(
      systemError("SetConversationStartersFailed", `Failed to set conversation starters: ${e}`, {
        source: "declarativeAgent/capabilities",
        inner: e instanceof Error ? e : new Error(String(e)),
      })
    );
  }
}

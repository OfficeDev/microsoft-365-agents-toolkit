// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as path from "path";
import { ok, err, Result } from "neverthrow";
import { getManifestPath, readTeamsManifest } from "../../manifest/readManifest";
import { userError, AtkError } from "../../core/error";

/**
 * Discover the Declarative Agent manifest path from a project folder.
 *
 * Algorithm:
 *  1. Find the Teams `manifest.json` via `getManifestPath(projectPath)`
 *  2. Read and parse it into a `TeamsManifestWrapper`
 *  3. Extract `copilotAgents.declarativeAgents[0].file`
 *  4. Resolve the relative path against the manifest directory
 *
 * @param projectPath  Absolute path to the project root
 * @returns Absolute path to the declarative agent manifest (e.g. `declarativeAgent.json`)
 */
export async function getAgentManifestPath(projectPath: string): Promise<Result<string, AtkError>> {
  const manifestPathResult = getManifestPath(projectPath);
  if (manifestPathResult.isErr()) return err(manifestPathResult.error);

  const teamsManifestPath = manifestPathResult.value;
  const readResult = await readTeamsManifest(teamsManifestPath);
  if (readResult.isErr()) return err(readResult.error);

  const wrapper = readResult.value.wrapper;
  const daRefs = wrapper.declarativeAgents;
  if (!daRefs || daRefs.length === 0) {
    return err(
      userError(
        "NoDeclarativeAgent",
        "No declarative agent found in manifest.json. Ensure copilotAgents.declarativeAgents is defined.",
        { source: "DeclarativeAgent" }
      )
    );
  }

  const daFile = (daRefs[0] as unknown as Record<string, unknown>).file as string | undefined;
  if (!daFile) {
    return err(
      userError(
        "NoDeclarativeAgentFile",
        "Declarative agent entry in manifest.json has no 'file' property.",
        { source: "DeclarativeAgent" }
      )
    );
  }

  const agentManifestPath = path.resolve(path.dirname(teamsManifestPath), daFile);
  return ok(agentManifestPath);
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs";
import * as path from "path";
import { Result, ok, err } from "neverthrow";
import { DeclarativeAgentManifestWrapper } from "@microsoft/app-manifest";
import type { AtkError } from "../../core/error";
import { systemError } from "../../core/error";

/** Default subdirectory where embedded knowledge files are stored. */
const EMBEDDED_KNOWLEDGE_DIR = "knowledge";

/**
 * Add embedded knowledge files to a declarative agent manifest.
 *
 * Each file in `sourceFilePaths` is copied into a `knowledge/` subdirectory
 * next to the manifest, and an entry is added to the EmbeddedKnowledge capability.
 *
 * @param agentManifestPath  Absolute path to the declarative agent manifest.
 * @param sourceFilePaths    Absolute paths of the files to embed.
 */
export async function addEmbeddedKnowledge(
  agentManifestPath: string,
  sourceFilePaths: string[]
): Promise<Result<void, AtkError>> {
  try {
    const manifestDir = path.dirname(agentManifestPath);
    const knowledgeDir = path.resolve(manifestDir, EMBEDDED_KNOWLEDGE_DIR);

    // Ensure knowledge directory exists
    await fs.promises.mkdir(knowledgeDir, { recursive: true });

    const wrapper = await DeclarativeAgentManifestWrapper.read(agentManifestPath);

    const relativeFiles: string[] = [];
    for (const srcPath of sourceFilePaths) {
      const destPath = path.resolve(knowledgeDir, path.basename(srcPath));
      await fs.promises.copyFile(srcPath, destPath);
      relativeFiles.push(path.relative(manifestDir, destPath).replace(/\\/g, "/"));
    }

    wrapper.addEmbeddedKnowledgeCapability(relativeFiles.map((f) => ({ file: f })));
    await wrapper.save(agentManifestPath);
    return ok(undefined);
  } catch (e) {
    return err(
      systemError("AddEmbeddedKnowledgeFailed", `Failed to add embedded knowledge: ${e}`, {
        source: "declarativeAgent/knowledge",
        inner: e instanceof Error ? e : new Error(String(e)),
      })
    );
  }
}

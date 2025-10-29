// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs-extra";
import { parseDocument, YAMLSeq } from "yaml";

/**
 * Merge source YAML file into target YAML file at stage level (provision, deploy, publish).
 * Rules:
 *  - For each top-level stage (provision, deploy, publish):
 *    - If stage doesn't exist in target, add it with all actions from source.
 *    - If stage exists in both, append source actions to target actions (array append).
 *  - Other top-level keys: if missing in target, add them; if present, keep target value.
 *  - Comments and formatting in target file are preserved where possible.
 */
export async function mergeYamlFile(sourcePath: string, targetPath: string): Promise<void> {
  if (!(await fs.pathExists(sourcePath))) return; // nothing to merge
  if (!(await fs.pathExists(targetPath))) {
    // If target does not exist, just copy source over
    await fs.copy(sourcePath, targetPath);
    return;
  }

  const rawSource = await fs.readFile(sourcePath, "utf8");
  const rawTarget = await fs.readFile(targetPath, "utf8");

  // Parse with yaml library to retain comments and formatting
  // Use strict: false and prettyErrors: false to be lenient with mustache syntax like {{appName}}
  const sourceDoc = parseDocument(rawSource, { strict: false, prettyErrors: false });
  const targetDoc = parseDocument(rawTarget, { strict: false, prettyErrors: false });

  const sourceContent = sourceDoc.toJSON();
  const targetContent = targetDoc.toJSON();

  if (!sourceContent || typeof sourceContent !== "object") return;
  if (!targetContent || typeof targetContent !== "object") {
    // Target is empty or malformed; use source
    await fs.copy(sourcePath, targetPath);
    return;
  }

  // Stages to merge (append actions arrays)
  const stages = ["provision", "deploy", "publish"];

  for (const stage of stages) {
    if (stage in sourceContent && Array.isArray(sourceContent[stage])) {
      if (stage in targetContent && Array.isArray(targetContent[stage])) {
        // Both have this stage; append source actions to target
        const targetStageNode = targetDoc.get(stage);
        const sourceStageNode = sourceDoc.get(stage);
        if (
          targetStageNode &&
          targetStageNode instanceof YAMLSeq &&
          sourceStageNode &&
          sourceStageNode instanceof YAMLSeq
        ) {
          // Get items from source YAMLSeq and add to target YAMLSeq
          for (let i = 0; i < sourceStageNode.items.length; i++) {
            targetStageNode.items.unshift(sourceStageNode.items[i]);
          }
        }
      } else {
        // Target doesn't have this stage; add it from source document
        const sourceStageNode = sourceDoc.get(stage);
        if (sourceStageNode) {
          targetDoc.set(stage, sourceStageNode);
        }
      }
    }
  }

  // Merge other top-level keys (not stages): add if missing, keep existing if present
  for (const key of Object.keys(sourceContent)) {
    if (!stages.includes(key) && !(key in targetContent)) {
      targetDoc.set(key, sourceContent[key]);
    }
  }

  let output = "";
  output = targetDoc.toString();
  await fs.writeFile(targetPath, output, "utf8");
}

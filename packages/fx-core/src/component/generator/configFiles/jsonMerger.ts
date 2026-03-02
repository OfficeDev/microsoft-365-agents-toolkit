// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as commentJson from "comment-json";
import * as fs from "fs-extra";

/**
 * Merge source JSON file (with comments) into target JSON file (with comments) at node level.
 * Rules:
 *  - If a top-level or nested property in source does not exist in target, add it (preserving source value).
 *  - If property exists and both values are arrays, append source array elements to target array (with de-duplication).
 *  - If property exists and both values are plain objects, recurse.
 *  - Otherwise (primitive / mismatched types), keep the existing target value (no overwrite).
 *  - Comments present in the existing target file are preserved. New properties won't have comments unless provided in target.
 */
export async function mergeJsonFile(sourcePath: string, targetPath: string): Promise<void> {
  if (!(await fs.pathExists(sourcePath))) return; // nothing to merge
  if (!(await fs.pathExists(targetPath))) {
    // If target does not exist just copy source over
    await fs.copy(sourcePath, targetPath);
    return;
  }

  const rawSource = await fs.readFile(sourcePath, "utf8");
  const rawTarget = await fs.readFile(targetPath, "utf8");

  // Parse with comment-json to retain comments/whitespace metadata where possible
  // Do NOT pass 'true' as third param (that would strip comments)
  const source = commentJson.parse(rawSource);
  const target = commentJson.parse(rawTarget);

  const merged = mergeNodes(target, source);

  const output = commentJson.stringify(merged, null, 2) + "\n";
  await fs.writeFile(targetPath, output, "utf8");
}

function isPlainObject(val: unknown): val is Record<string, unknown> {
  return !!val && typeof val === "object" && !Array.isArray(val);
}

function arrayContains(arr: unknown[], el: unknown): boolean {
  // For primitives, use strict equality
  if (typeof el !== "object" || el === null) {
    return arr.includes(el);
  }
  // For objects/arrays, use deep equality check
  return arr.some((item) => JSON.stringify(item) === JSON.stringify(el));
}

function mergeNodes(
  target: commentJson.CommentJSONValue,
  source: commentJson.CommentJSONValue
): commentJson.CommentJSONValue {
  if (Array.isArray(target) && Array.isArray(source)) {
    // Append all elements from source (retain order, no duplicates)
    for (const el of source) {
      if (!arrayContains(target, el)) {
        (target as commentJson.CommentArray<commentJson.CommentJSONValue>).push(el);
      }
    }
    return target;
  }
  if (isPlainObject(target) && isPlainObject(source)) {
    for (const key of Object.keys(source)) {
      const sVal = (source as Record<string, commentJson.CommentJSONValue>)[key];
      if (!(key in (target as Record<string, commentJson.CommentJSONValue>))) {
        (target as Record<string, commentJson.CommentJSONValue>)[key] = sVal;
      } else {
        const tVal = (target as Record<string, commentJson.CommentJSONValue>)[key];
        if (Array.isArray(tVal) && Array.isArray(sVal)) {
          for (const el of sVal) {
            if (!arrayContains(tVal, el)) {
              (tVal as commentJson.CommentArray<commentJson.CommentJSONValue>).push(el);
            }
          }
        } else if (isPlainObject(tVal) && isPlainObject(sVal)) {
          mergeNodes(
            tVal as unknown as commentJson.CommentJSONValue,
            sVal as unknown as commentJson.CommentJSONValue
          );
        } // else keep existing primitive / mismatched types
      }
    }
    return target;
  }
  // For primitives or mismatched types: do not overwrite target per requirements
  return target;
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "node:fs/promises";
import * as path from "node:path";
import { ok, err } from "neverthrow";
import { z } from "zod";
import { createDriver } from "../../createDriver";
import { userError } from "../../../core/error";

const inputSchema = z
  .object({
    /** Relative or absolute path to the JSON file */
    target: z.string().min(1, "target must be a non-empty path"),
    /** Content object to deep-merge into the file (preferred) */
    content: z.record(z.unknown()).optional(),
    /** Legacy alias for content (used by older templates) */
    appsettings: z.record(z.unknown()).optional(),
  })
  .refine((data) => data.content || data.appsettings, {
    message: "Either 'content' or 'appsettings' must be provided",
  });

type CreateOrUpdateJsonFileConfig = z.infer<typeof inputSchema>;

/**
 * Driver: file/createOrUpdateJsonFile
 *
 * Creates or updates a JSON file by deep-merging the provided content.
 * Existing keys at nested levels are preserved unless overridden.
 */
export const createOrUpdateJsonFileDriver = createDriver<CreateOrUpdateJsonFileConfig>({
  id: "file/createOrUpdateJsonFile",
  name: "Create or Update JSON File",
  inputSchema,
  execute: async (ctx, config) => {
    const projectPath = ctx.projectPath;
    if (!projectPath) {
      return err(
        userError("ProjectPathRequired", "projectPath is required to resolve file target.", {
          source: "file/createOrUpdateJsonFile",
        })
      );
    }

    const targetPath = path.isAbsolute(config.target)
      ? config.target
      : path.resolve(projectPath, config.target);

    // Ensure the directory exists
    await fs.mkdir(path.dirname(targetPath), { recursive: true });

    // Read existing JSON or start fresh
    let existing: Record<string, unknown> = {};
    try {
      const raw = await fs.readFile(targetPath, "utf-8");
      existing = JSON.parse(raw);
    } catch {
      // File doesn't exist or isn't valid JSON — start from empty
    }

    const payload = config.content ?? config.appsettings ?? {};
    const merged = deepMerge(existing, payload);

    await fs.writeFile(targetPath, JSON.stringify(merged, null, "\t") + "\n", "utf-8");

    ctx.logger.info(`[file/createOrUpdateJsonFile] Wrote merged JSON to ${targetPath}`);

    return ok({ outputs: {} });
  },
});

/**
 * Deep-merge `source` into `target`.
 * Only plain objects are recursed; arrays and primitives are replaced.
 */
function deepMerge(
  target: Record<string, unknown>,
  source: Record<string, unknown>
): Record<string, unknown> {
  const result = { ...target };
  for (const key of Object.keys(source)) {
    const srcVal = source[key];
    const tgtVal = result[key];
    if (isPlainObject(srcVal) && isPlainObject(tgtVal)) {
      result[key] = deepMerge(tgtVal, srcVal);
    } else {
      result[key] = srcVal;
    }
  }
  return result;
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

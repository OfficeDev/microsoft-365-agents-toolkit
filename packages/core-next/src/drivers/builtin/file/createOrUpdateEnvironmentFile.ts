// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "node:fs/promises";
import * as path from "node:path";
import * as os from "node:os";
import { ok, err } from "neverthrow";
import { z } from "zod";
import { createDriver } from "../../createDriver";
import { userError } from "../../../core/error";

const inputSchema = z.object({
  /** Relative or absolute path to the .env file */
  target: z.string().min(1, "target must be a non-empty path"),
  /** Key-value pairs to write/merge into the file */
  envs: z.record(z.string()),
});

type CreateOrUpdateEnvFileConfig = z.infer<typeof inputSchema>;

/**
 * Driver: file/createOrUpdateEnvironmentFile
 *
 * Creates or updates a `.env` file by merging key-value pairs.
 * Existing keys are preserved unless overridden by the new values.
 * No process.env side effects — outputs are returned for the lifecycle executor.
 */
export const createOrUpdateEnvironmentFileDriver = createDriver<CreateOrUpdateEnvFileConfig>({
  id: "file/createOrUpdateEnvironmentFile",
  name: "Create or Update Environment File",
  inputSchema,
  execute: async (ctx, config) => {
    const projectPath = ctx.projectPath;
    if (!projectPath) {
      return err(
        userError("ProjectPathRequired", "projectPath is required to resolve file target.", {
          source: "file/createOrUpdateEnvironmentFile",
        })
      );
    }

    const targetPath = path.isAbsolute(config.target)
      ? config.target
      : path.resolve(projectPath, config.target);

    // Ensure the directory exists
    await fs.mkdir(path.dirname(targetPath), { recursive: true });

    // Read existing env vars if file exists
    let existing: Record<string, string> = {};
    try {
      const content = await fs.readFile(targetPath, "utf-8");
      existing = parseEnvFile(content);
    } catch {
      // File doesn't exist yet — will create it
    }

    // Merge: new values override existing
    const merged = { ...existing, ...config.envs };

    // Format and write
    const lines = Object.entries(merged).map(([key, value]) => `${key}=${value}`);
    await fs.writeFile(targetPath, lines.join(os.EOL) + os.EOL, "utf-8");

    ctx.logger.info(
      `[file/createOrUpdateEnvironmentFile] Wrote ${Object.keys(config.envs).length} env var(s) to ${targetPath}`
    );

    return ok({
      outputs: config.envs,
    });
  },
});

/**
 * Parse a .env file into key-value pairs.
 * Handles `KEY=VALUE`, ignores comments (#) and blank lines.
 */
function parseEnvFile(content: string): Record<string, string> {
  const result: Record<string, string> = {};
  for (const line of content.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const eqIndex = trimmed.indexOf("=");
    if (eqIndex === -1) continue;
    const key = trimmed.slice(0, eqIndex).trim();
    let value = trimmed.slice(eqIndex + 1);
    // Strip surrounding quotes if present
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }
    result[key] = value;
  }
  return result;
}

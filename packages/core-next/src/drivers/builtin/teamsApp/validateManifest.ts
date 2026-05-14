// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as path from "node:path";
import { z } from "zod";
import { ok, err } from "neverthrow";
import { createDriver } from "../../createDriver";
import { validateManifestSchema } from "../../../manifest/validate";

const inputSchema = z.object({
  /** Path to the Teams app manifest (manifest.json) */
  manifestPath: z.string().min(1),
});

/**
 * Driver: teamsApp/validateManifest
 *
 * Validates a Teams app manifest against the JSON schema.
 *
 * Outputs:
 * - TEAMS_APP_MANIFEST_VALID: "true" or "false"
 */
export const validateManifestDriver = createDriver({
  id: "teamsApp/validateManifest",
  name: "Validate Manifest",
  inputSchema,
  execute: async (ctx, config) => {
    const manifestPath = path.isAbsolute(config.manifestPath)
      ? config.manifestPath
      : path.resolve(ctx.projectPath ?? process.cwd(), config.manifestPath);
    const result = await validateManifestSchema(manifestPath);
    if (result.isErr()) return err(result.error);

    const { errors } = result.value;
    const valid = errors.length === 0;

    return ok({
      outputs: { TEAMS_APP_MANIFEST_VALID: String(valid) },
    });
  },
});

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ok, err, Result } from "neverthrow";
import { TeamsManifestWrapper } from "@microsoft/app-manifest";
import { systemError, AtkError } from "../core/error";

/**
 * Validate a Teams app manifest.json against its JSON schema.
 *
 * Delegates to `TeamsManifestWrapper.validate()` which performs full
 * JSON-schema validation using the bundled schema for the declared
 * `$schema` / `manifestVersion`.
 *
 * @param manifestPath  Absolute path to manifest.json
 * @returns List of validation errors (empty = valid), or an AtkError
 */
export async function validateManifestSchema(
  manifestPath: string
): Promise<Result<{ valid: boolean; errors: string[] }, AtkError>> {
  try {
    const wrapper = await TeamsManifestWrapper.read(manifestPath);
    const errors = await wrapper.validate();
    return ok({ valid: errors.length === 0, errors });
  } catch (e) {
    return err(
      systemError(
        "ValidateManifestFailed",
        `Failed to validate manifest at ${manifestPath}: ${e}`,
        {
          source: "ManifestValidate",
          inner: e instanceof Error ? e : new Error(String(e)),
        }
      )
    );
  }
}

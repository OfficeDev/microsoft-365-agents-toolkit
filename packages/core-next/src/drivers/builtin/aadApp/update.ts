// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { z } from "zod";
import { ok, err } from "neverthrow";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import { createDriver } from "../../createDriver";
import { systemError, userError } from "../../../core/error";
import { GraphApiClient } from "../../../clients/graphApi/client";
import {
  graphScopes,
  AADApplication,
  RequiredResourceAccess,
} from "../../../clients/graphApi/types";
import { resolveEnvPlaceholders } from "../../../manifest/resolve";
import * as permissionListJson from "./permissions.json";

/**
 * UUID v4 regex — used to detect whether a string is already a GUID.
 */
const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

interface PermissionEntry {
  id: string; // e.g. "f431331c-…"
  value: string; // e.g. "ExternalConnection.ReadWrite.OwnedBy"
}

interface ServicePrincipalJson {
  appId: string;
  displayName: string;
  appRoles: PermissionEntry[];
  oauth2PermissionScopes: PermissionEntry[];
}

interface PermissionMap {
  [appIdOrName: string]: {
    id: string;
    roles: Record<string, string>; // friendly-name → GUID
    scopes: Record<string, string>; // friendly-name → GUID
  };
}

let cachedPermissionMap: PermissionMap | null = null;

/**
 * Build a map from service-principal display-names and appIds to their
 * role/scope name→GUID mappings.  Matches fx-core's getPermissionMap().
 */
function getPermissionMap(): PermissionMap {
  if (cachedPermissionMap) return cachedPermissionMap;
  const list = permissionListJson as { value: ServicePrincipalJson[] };
  const map: PermissionMap = {};
  for (const sp of list.value) {
    const entry = {
      id: sp.appId,
      roles: {} as Record<string, string>,
      scopes: {} as Record<string, string>,
    };
    for (const r of sp.appRoles) {
      entry.roles[r.value] = r.id;
    }
    for (const s of sp.oauth2PermissionScopes) {
      entry.scopes[s.value] = s.id;
    }
    map[sp.appId] = entry;
    map[sp.displayName] = entry;
  }
  cachedPermissionMap = map;
  return map;
}

/**
 * Resolve friendly permission names (e.g. "Microsoft Graph",
 * "ExternalConnection.ReadWrite.OwnedBy") to their GUID equivalents
 * in-place, matching fx-core's processRequiredResourceAccessInManifest().
 */
function resolvePermissionNames(rra: RequiredResourceAccess[]): void {
  const map = getPermissionMap();
  for (const item of rra) {
    // Resolve resourceAppId (e.g. "Microsoft Graph" → "00000003-…")
    if (!UUID_RE.test(item.resourceAppId)) {
      const entry = map[item.resourceAppId];
      if (entry) {
        item.resourceAppId = entry.id;
      }
    }
    // Resolve each resourceAccess.id
    const entry = map[item.resourceAppId];
    if (!entry) continue;
    for (const ra of item.resourceAccess) {
      if (!UUID_RE.test(ra.id)) {
        const resolved = ra.type === "Scope" ? entry.scopes[ra.id] : entry.roles[ra.id];
        if (resolved) {
          ra.id = resolved;
        }
      }
    }
  }
}

const inputSchema = z.object({
  /**
   * Path to the AAD manifest JSON file (relative to projectPath).
   * Supports `${{ ENV_VAR }}` placeholders that are resolved by the lifecycle executor.
   */
  manifestPath: z.string().min(1),
  /** Path where the resolved manifest will be written (relative to projectPath). */
  outputFilePath: z.string().min(1),
});

/**
 * Driver: aadApp/update
 *
 * Reads an Entra ID manifest template, resolves env placeholders, and
 * PATCHes the application in Graph API.
 *
 * Outputs:
 * - AAD_APP_ACCESS_AS_USER_PERMISSION_ID (auto-generated UUID if referenced but missing)
 */
export const updateAadAppDriver = createDriver({
  id: "aadApp/update",
  name: "Update AAD App",
  inputSchema,
  execute: async (ctx, config) => {
    const source = "aadApp/update";

    if (!ctx.projectPath) {
      return err(
        userError("MissingProjectPath", "projectPath is required for aadApp/update", { source })
      );
    }

    // Resolve file paths
    const manifestFullPath = path.resolve(ctx.projectPath, config.manifestPath);
    const outputFullPath = path.resolve(ctx.projectPath, config.outputFilePath);

    // Read manifest template
    let manifestContent: string;
    try {
      manifestContent = await fs.readFile(manifestFullPath, "utf-8");
    } catch {
      return err(
        userError("ManifestFileNotFound", `AAD manifest not found: ${manifestFullPath}`, { source })
      );
    }

    // Auto-generate AAD_APP_ACCESS_AS_USER_PERMISSION_ID if the manifest references
    // it but no value exists in the environment yet.  Matches fx-core behaviour:
    // the UUID is created once and persisted via the driver's outputs.
    const permissionIdPlaceholder = /\$\{\{ *AAD_APP_ACCESS_AS_USER_PERMISSION_ID *\}\}/;
    let generatedPermissionId: string | undefined;
    if (
      !process.env.AAD_APP_ACCESS_AS_USER_PERMISSION_ID &&
      permissionIdPlaceholder.test(manifestContent)
    ) {
      const { randomUUID } = await import("node:crypto");
      generatedPermissionId = randomUUID();
      process.env.AAD_APP_ACCESS_AS_USER_PERMISSION_ID = generatedPermissionId;
    }

    // Resolve ${{VAR}} env placeholders before parsing JSON
    const { content: resolvedContent, unresolved } = resolveEnvPlaceholders(manifestContent);
    if (unresolved.length > 0) {
      return err(
        userError(
          "UnresolvedManifestVars",
          `AAD manifest has unresolved env variables: ${unresolved.map((u) => u.name).join(", ")}`,
          { source }
        )
      );
    }

    // Parse manifest
    let manifest: AADApplication;
    try {
      manifest = JSON.parse(resolvedContent);
    } catch {
      return err(
        userError("InvalidManifestJson", `Invalid JSON in AAD manifest: ${manifestFullPath}`, {
          source,
        })
      );
    }

    // Validate that the manifest has an object ID
    const objectId = manifest.id;
    if (!objectId) {
      return err(
        userError("MissingObjectId", "AAD manifest must have an 'id' (object ID) field", {
          source,
        })
      );
    }

    // Resolve friendly permission names to GUIDs (e.g. "Microsoft Graph" → appId,
    // "ExternalConnection.ReadWrite.OwnedBy" → role GUID).  Matches fx-core behaviour.
    if (manifest.requiredResourceAccess?.length) {
      resolvePermissionNames(manifest.requiredResourceAccess);
    }

    // Write resolved manifest to output path
    await fs.mkdir(path.dirname(outputFullPath), { recursive: true });
    await fs.writeFile(outputFullPath, JSON.stringify(manifest, null, 2), "utf-8");

    // Acquire Graph token
    const tokenRes = await ctx.auth.m365TokenProvider.getAccessToken({
      scopes: graphScopes(),
    });
    if (tokenRes.isErr()) {
      return err(
        systemError("TokenAcquisitionError", tokenRes.error.message, {
          source,
          inner: tokenRes.error,
        })
      );
    }
    const client = new GraphApiClient(ctx, tokenRes.value);

    // Build update payload — strip read-only fields
    const updates: Partial<AADApplication> = { ...manifest };
    delete updates.id; // object ID is not patchable
    delete updates.appId; // read-only

    // Two-phase update for preAuthorizedApplications (Graph requires permissions to exist first)
    if (updates.api?.preAuthorizedApplications?.length) {
      const preAuth = updates.api.preAuthorizedApplications;
      const phase1 = {
        ...updates,
        api: { ...updates.api, preAuthorizedApplications: [] },
      };
      const res1 = await client.updateAadApp(objectId, phase1);
      if (res1.isErr()) return err(res1.error);

      const phase2 = { api: { preAuthorizedApplications: preAuth } };
      const res2 = await client.updateAadApp(objectId, phase2);
      if (res2.isErr()) return err(res2.error);
    } else {
      const res = await client.updateAadApp(objectId, updates);
      if (res.isErr()) return err(res.error);
    }

    ctx.logger.info(`[${source}] Updated AAD app: objectId=${objectId}`);

    // Output the generated permission ID so the executor persists it to envMap.
    // On subsequent runs the value comes from the env file and this path is skipped.
    const outputs: Record<string, string> = {};
    if (generatedPermissionId) {
      outputs.AAD_APP_ACCESS_AS_USER_PERMISSION_ID = generatedPermissionId;
    }

    return ok({ outputs });
  },
});

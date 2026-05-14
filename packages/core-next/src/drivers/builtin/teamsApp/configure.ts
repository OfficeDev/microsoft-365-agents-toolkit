// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { z } from "zod";
import { ok, err } from "neverthrow";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import AdmZip from "adm-zip";
import { createDriver } from "../../createDriver";
import { userError, systemError } from "../../../core/error";
import { TeamsDevPortalClient } from "../../../clients/teamsDevPortal/client";
import { appStudioScopes } from "../../../clients/teamsDevPortal/types";

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const inputSchema = z.object({
  /** Path to the zipped app package (.zip) */
  appPackagePath: z.string().min(1),
});

/**
 * Driver: teamsApp/configure
 *
 * Updates an existing Teams app in the Developer Portal.
 * Reads the manifest from the provided ZIP, verifies the app exists,
 * then uploads the package with overwrite=true.
 *
 * Outputs:
 * - TEAMS_APP_TENANT_ID: tenant that owns the app
 * - TEAMS_APP_UPDATE_TIME: timestamp of the update
 */
export const configureTeamsAppDriver = createDriver({
  id: "teamsApp/configure",
  name: "Configure Teams App",
  inputSchema,
  execute: async (ctx, config) => {
    // Resolve relative paths against project root (YAML paths are project-relative).
    // ATK convention: CWD is the project root; fall back to process.cwd().
    const absPackagePath = path.isAbsolute(config.appPackagePath)
      ? config.appPackagePath
      : path.resolve(ctx.projectPath ?? process.cwd(), config.appPackagePath);

    // Read the app package directly (no separate existence check to avoid TOCTOU race)
    let archivedFile: Buffer;
    try {
      archivedFile = Buffer.from(await fs.readFile(absPackagePath));
    } catch (e: unknown) {
      if ((e as NodeJS.ErrnoException).code === "ENOENT") {
        return err(
          userError("AppPackageNotFound", `App package not found: ${config.appPackagePath}`, {
            source: "teamsApp/configure",
            help: "https://aka.ms/teamsfx-actions/teamsapp-update",
          })
        );
      }
      throw e;
    }

    // Extract manifest to read teamsAppId
    const zip = new AdmZip(archivedFile);
    const manifestEntry = zip.getEntry("manifest.json");
    if (!manifestEntry) {
      return err(
        userError("ManifestNotFound", "manifest.json not found inside app package", {
          source: "teamsApp/configure",
        })
      );
    }
    const manifest = JSON.parse(manifestEntry.getData().toString("utf8")) as { id?: string };
    const teamsAppId = manifest.id;
    if (!teamsAppId || !UUID_RE.test(teamsAppId)) {
      return err(
        userError("InvalidTeamsAppId", `Invalid Teams App ID in manifest: "${teamsAppId ?? ""}"`, {
          source: "teamsApp/configure",
        })
      );
    }

    // Acquire M365 token
    const tokenRes = await ctx.auth.m365TokenProvider.getAccessToken({
      scopes: appStudioScopes(),
    });
    if (tokenRes.isErr()) {
      return err(
        systemError("TokenAcquisitionError", tokenRes.error.message, {
          source: "teamsApp/configure",
          inner: tokenRes.error,
        })
      );
    }
    const client = new TeamsDevPortalClient(ctx, tokenRes.value);

    // Verify the app exists
    const existCheck = await client.getApp(teamsAppId);
    if (existCheck.isErr()) {
      return err(
        userError(
          "TeamsAppNotFound",
          `App ${teamsAppId} not found in Developer Portal — cannot configure. Create it first.`,
          { source: "teamsApp/configure" }
        )
      );
    }

    // Upload with overwrite
    const importRes = await client.importApp(archivedFile, true);
    if (importRes.isErr()) return err(importRes.error);

    const updated = importRes.value;
    ctx.logger.info(`[teamsApp/configure] Updated app ${teamsAppId}`);

    return ok({
      outputs: {
        TEAMS_APP_TENANT_ID: updated.tenantId ?? "",
        TEAMS_APP_UPDATE_TIME: updated.updatedAt ?? new Date().toISOString(),
      },
    });
  },
});

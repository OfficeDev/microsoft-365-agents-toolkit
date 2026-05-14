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
  appPackagePath: z.string().min(1),
});

/**
 * Driver: teamsApp/update
 *
 * Alias of teamsApp/configure — updates an existing Teams app in the Developer Portal.
 * Many template YAML files reference `teamsApp/update` rather than `teamsApp/configure`.
 */
export const updateTeamsAppDriver = createDriver({
  id: "teamsApp/update",
  name: "Update Teams App",
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
            source: "teamsApp/update",
            help: "https://aka.ms/teamsfx-actions/teamsapp-update",
          })
        );
      }
      throw e;
    }

    const zip = new AdmZip(archivedFile);
    const manifestEntry = zip.getEntry("manifest.json");
    if (!manifestEntry) {
      return err(
        userError("ManifestNotFound", "manifest.json not found inside app package", {
          source: "teamsApp/update",
        })
      );
    }
    const manifest = JSON.parse(manifestEntry.getData().toString("utf8")) as { id?: string };
    const teamsAppId = manifest.id;
    if (!teamsAppId || !UUID_RE.test(teamsAppId)) {
      return err(
        userError("InvalidTeamsAppId", `Invalid Teams App ID in manifest: "${teamsAppId ?? ""}"`, {
          source: "teamsApp/update",
        })
      );
    }

    const tokenRes = await ctx.auth.m365TokenProvider.getAccessToken({
      scopes: appStudioScopes(),
    });
    if (tokenRes.isErr()) {
      return err(
        systemError("TokenAcquisitionError", tokenRes.error.message, {
          source: "teamsApp/update",
          inner: tokenRes.error,
        })
      );
    }
    const client = new TeamsDevPortalClient(ctx, tokenRes.value);

    const existCheck = await client.getApp(teamsAppId);
    if (existCheck.isErr()) {
      return err(
        userError(
          "TeamsAppNotFound",
          `App ${teamsAppId} not found in Developer Portal — cannot update. Create it first.`,
          { source: "teamsApp/update" }
        )
      );
    }

    const importRes = await client.importApp(archivedFile, true);
    if (importRes.isErr()) return err(importRes.error);

    const updated = importRes.value;
    ctx.logger.info(`[teamsApp/update] Updated app ${teamsAppId}`);

    return ok({
      outputs: {
        TEAMS_APP_TENANT_ID: updated.tenantId ?? "",
        TEAMS_APP_UPDATE_TIME: updated.updatedAt ?? new Date().toISOString(),
      },
    });
  },
});

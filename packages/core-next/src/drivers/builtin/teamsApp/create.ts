// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { z } from "zod";
import { ok, err } from "neverthrow";
import { randomUUID } from "node:crypto";
import AdmZip from "adm-zip";
import { createDriver } from "../../createDriver";
import { systemError } from "../../../core/error";
import { TeamsDevPortalClient } from "../../../clients/teamsDevPortal/client";
import { appStudioScopes } from "../../../clients/teamsDevPortal/types";

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const inputSchema = z.object({
  /** Short name for the Teams app */
  name: z.string().min(1),
  /**
   * Existing Teams App ID for idempotency.
   * When the lifecycle executor resolves ${{TEAMS_APP_ID}} from the env map,
   * this will be a UUID if the app was already created. If the placeholder was
   * unresolved, it stays as the literal "${{TEAMS_APP_ID}}" and we skip the check.
   */
  existingTeamsAppId: z.string().optional(),
});

/** Minimal 1×1 transparent PNG (68 bytes) for the default icons */
const MINIMAL_PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVQI12NgAAIABQAB" +
    "Nl7BcQAAAABJRU5ErkJggg==",
  "base64"
);

/**
 * Driver: teamsApp/create
 *
 * Creates a new Teams app in the Developer Portal (or returns the existing one
 * if `existingTeamsAppId` points to a valid app — idempotent).
 *
 * Outputs:
 * - TEAMS_APP_ID: the Teams App ID
 * - TEAMS_APP_TENANT_ID: the tenant that owns the app
 */
export const createTeamsAppDriver = createDriver({
  id: "teamsApp/create",
  name: "Create Teams App",
  inputSchema,
  execute: async (ctx, config) => {
    // Acquire M365 token
    const tokenRes = await ctx.auth.m365TokenProvider.getAccessToken({
      scopes: appStudioScopes(),
    });
    if (tokenRes.isErr()) {
      return err(
        systemError("TokenAcquisitionError", tokenRes.error.message, {
          source: "teamsApp/create",
          inner: tokenRes.error,
        })
      );
    }
    const client = new TeamsDevPortalClient(ctx, tokenRes.value);

    // Idempotency: if we already have a valid app ID, check that it still exists
    if (config.existingTeamsAppId && UUID_RE.test(config.existingTeamsAppId)) {
      const existing = await client.getApp(config.existingTeamsAppId);
      if (existing.isOk()) {
        ctx.logger.info(
          `[teamsApp/create] App ${config.existingTeamsAppId} already exists — skipping creation`
        );
        return ok({
          outputs: {
            TEAMS_APP_ID: existing.value.teamsAppId!,
            TEAMS_APP_TENANT_ID: existing.value.tenantId ?? "",
          },
        });
      }
      // If getApp failed (404, etc.), fall through to create
    }

    // Build minimal app package
    const appId =
      config.existingTeamsAppId && UUID_RE.test(config.existingTeamsAppId)
        ? config.existingTeamsAppId
        : randomUUID();

    const manifest = {
      $schema:
        "https://developer.microsoft.com/en-us/json-schemas/teams/v1.17/MicrosoftTeams.schema.json",
      manifestVersion: "1.17",
      version: "1.0.0",
      id: appId,
      developer: {
        name: "Teams App Developer",
        websiteUrl: "https://example.com",
        privacyUrl: "https://example.com/privacy",
        termsOfUseUrl: "https://example.com/terms",
      },
      name: { short: config.name, full: config.name },
      description: {
        short: config.name,
        full: `${config.name} created by M365 Agents Toolkit`,
      },
      icons: { color: "color.png", outline: "outline.png" },
      accentColor: "#FFFFFF",
    };

    const zip = new AdmZip();
    zip.addFile("manifest.json", Buffer.from(JSON.stringify(manifest, null, 2)));
    zip.addFile("color.png", MINIMAL_PNG);
    zip.addFile("outline.png", MINIMAL_PNG);

    const importRes = await client.importApp(zip.toBuffer(), false);
    if (importRes.isErr()) {
      return err(importRes.error);
    }

    const created = importRes.value;
    ctx.logger.info(`[teamsApp/create] Created app ${created.teamsAppId}`);

    return ok({
      outputs: {
        TEAMS_APP_ID: created.teamsAppId!,
        TEAMS_APP_TENANT_ID: created.tenantId ?? "",
      },
    });
  },
});

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { z } from "zod";
import { ok, err } from "neverthrow";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import { createDriver } from "../../createDriver";
import { userError, systemError } from "../../../core/error";
import { M365PackageService, mosServiceScopes, AppScope } from "../../../clients/m365";

/** Map lowercase scope strings to AppScope enum values. */
function normalizeAppScope(val: string | undefined): AppScope | undefined {
  if (!val) return undefined;
  const lower = val.toLowerCase();
  const map: Record<string, AppScope> = {
    personal: AppScope.Personal,
    shared: AppScope.Shared,
    tenant: AppScope.Tenant,
  };
  return map[lower];
}

const inputSchema = z.object({
  /** Path to the zipped app package (.zip) */
  appPackagePath: z.string().min(1),
  /** Scope for sideloading: Personal, Shared, or Tenant (case-insensitive) */
  scope: z.string().optional(),
});

/**
 * Driver: teamsApp/extendToM365
 *
 * Sideloads a Teams app package to the M365 ecosystem (Outlook, Microsoft 365 app)
 * via the MOS (Microsoft Orchestrator Service) PackageService API.
 *
 * Outputs:
 * - titleId → M365_TITLE_ID
 * - appId  → M365_APP_ID
 * - shareLink → SHARE_LINK (optional, only for Shared scope)
 */
export const extendToM365Driver = createDriver({
  id: "teamsApp/extendToM365",
  name: "Extend Teams App to M365",
  inputSchema,
  execute: async (ctx, config) => {
    const packagePath = path.isAbsolute(config.appPackagePath)
      ? config.appPackagePath
      : path.resolve(ctx.projectPath ?? ".", config.appPackagePath);

    if (
      !(await fs.access(packagePath).then(
        () => true,
        () => false
      ))
    ) {
      return err(
        userError("AppPackageNotFound", `App package not found: ${packagePath}`, {
          source: "teamsApp/extendToM365",
          help: "https://aka.ms/teamsfx-actions/teamsapp-extendToM365",
        })
      );
    }

    // Acquire M365 sideloading token
    const tokenRes = await ctx.auth.m365TokenProvider.getAccessToken({
      scopes: mosServiceScopes(),
    });
    if (tokenRes.isErr()) {
      return err(
        systemError("TokenAcquisitionError", tokenRes.error.message, {
          source: "teamsApp/extendToM365",
          inner: tokenRes.error,
        })
      );
    }

    const service = new M365PackageService(ctx, tokenRes.value);
    const scope = normalizeAppScope(config.scope) ?? AppScope.Personal;

    ctx.logger.info(`[teamsApp/extendToM365] Sideloading to M365 (scope: ${scope})...`);

    const sideloadRes = await service.sideLoad(packagePath, scope);
    if (sideloadRes.isErr()) return err(sideloadRes.error);

    const { titleId, appId, shareLink } = sideloadRes.value;

    ctx.logger.info(`[teamsApp/extendToM365] Sideloaded: titleId=${titleId}, appId=${appId}`);

    const outputs: Record<string, string> = {
      titleId,
      appId,
    };
    if (shareLink) {
      outputs.shareLink = shareLink;
    }

    return ok({ outputs });
  },
});

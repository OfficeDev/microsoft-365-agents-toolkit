// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs";
import * as path from "path";
import AdmZip from "adm-zip";
import { z } from "zod";
import { ok, err } from "neverthrow";
import { TeamsManifestWrapper } from "@microsoft/app-manifest";
import { defineOperation } from "../core/operation";
import type { AtkContext } from "../core/context";
import { userError, systemError } from "../core/error";
import { buildAppPackage } from "./packageBuilder";
import { M365PackageService, mosServiceScopes, AppScope } from "../clients/m365";

// ─── Validate Manifest ───────────────────────────────────────

const validateInputSchema = z.object({
  /** Path to the Teams app manifest (manifest.json) */
  manifestPath: z.string().min(1),
});

interface ValidateResult {
  valid: boolean;
  errors: string[];
  warnings: string[];
}

/**
 * Validate a Teams app manifest against the JSON schema.
 */
export const validateManifestOp = defineOperation(
  "teamsApp/validate",
  validateInputSchema,
  async (_ctx: AtkContext, input: z.infer<typeof validateInputSchema>) => {
    try {
      const wrapper = await TeamsManifestWrapper.read(input.manifestPath);
      const errors = await wrapper.validate();

      return ok({
        valid: errors.length === 0,
        errors,
        warnings: [],
      } satisfies ValidateResult);
    } catch (e) {
      return err(
        systemError("ValidateManifestFailed", `Failed to validate manifest: ${e}`, {
          source: "teamsApp/validate",
          inner: e instanceof Error ? e : new Error(String(e)),
        })
      );
    }
  }
);

// ─── Package ─────────────────────────────────────────────────

const packageInputSchema = z.object({
  /** Project root directory */
  projectPath: z.string().min(1),
  /** Optional explicit manifest path (auto-detected if omitted) */
  manifestPath: z.string().optional(),
  /** Output path for the .zip app package */
  outputPath: z.string().min(1),
  /** Optional folder for resolved manifest files */
  outputFolder: z.string().optional(),
});

/**
 * Package a Teams app manifest directory into a .zip app package.
 * Delegates to buildAppPackage() for full env-var resolution,
 * DA/plugin handling, and icon processing.
 */
export const packageAppOp = defineOperation(
  "teamsApp/package",
  packageInputSchema,
  async (ctx: AtkContext, input: z.infer<typeof packageInputSchema>) => {
    return buildAppPackage(ctx, {
      projectPath: input.projectPath,
      manifestPath: input.manifestPath,
      outputZipPath: input.outputPath,
      outputFolder: input.outputFolder,
    });
  }
);

// ─── Validate App Package ────────────────────────────────────

const validatePackageInputSchema = z.object({
  /** Path to the .zip app package to validate */
  packagePath: z.string().min(1),
});

interface ValidatePackageResult {
  valid: boolean;
  errors: string[];
}

/**
 * Validate a built Teams app package (.zip).
 * Opens the ZIP, reads manifest.json, and validates it.
 */
export const validateAppPackageOp = defineOperation(
  "teamsApp/validateAppPackage",
  validatePackageInputSchema,
  async (_ctx: AtkContext, input: z.infer<typeof validatePackageInputSchema>) => {
    try {
      await fs.promises.access(input.packagePath);
    } catch {
      return err(
        userError("PackageNotFound", `App package not found: ${input.packagePath}`, {
          source: "teamsApp/validateAppPackage",
        })
      );
    }

    try {
      const zip = new AdmZip(input.packagePath);
      const manifestEntry = zip.getEntry("manifest.json");
      if (!manifestEntry) {
        return ok({
          valid: false,
          errors: ["manifest.json not found in app package"],
        } satisfies ValidatePackageResult);
      }

      const manifestContent = manifestEntry.getData().toString("utf8");
      const wrapper = TeamsManifestWrapper.fromJSON(manifestContent);
      const errors = await wrapper.validate();

      // Check icons exist in the ZIP
      const iconErrors: string[] = [];
      for (const iconName of [wrapper.icons.color, wrapper.icons.outline]) {
        if (!zip.getEntry(iconName)) {
          iconErrors.push(`Icon file "${iconName}" not found in app package`);
        }
      }

      const allErrors = [...errors, ...iconErrors];
      return ok({
        valid: allErrors.length === 0,
        errors: allErrors,
      } satisfies ValidatePackageResult);
    } catch (e) {
      return err(
        systemError("ValidateAppPackageFailed", `Failed to validate app package: ${e}`, {
          source: "teamsApp/validateAppPackage",
          inner: e instanceof Error ? e : new Error(String(e)),
        })
      );
    }
  }
);

// ─── Publish ─────────────────────────────────────────────────

const publishInputSchema = z.object({
  /** Path to the .zip app package to publish */
  packagePath: z.string().min(1),
});

/**
 * Publish a Teams app package.
 *
 * This is a placeholder that establishes the operation contract.
 * The actual API call requires Graph API integration which will be
 * implemented in the driver layer (teams-app-publish driver).
 */
export const publishAppOp = defineOperation(
  "teamsApp/publish",
  publishInputSchema,
  async (ctx: AtkContext, input: z.infer<typeof publishInputSchema>) => {
    // Verify package exists
    try {
      await fs.promises.access(input.packagePath);
    } catch {
      return err(
        userError("PackageNotFound", `App package not found: ${input.packagePath}`, {
          source: "teamsApp/publish",
        })
      );
    }

    // The actual publish is delegated to the teams-app-publish driver
    // via the lifecycle engine. This operation validates the contract.
    return err(
      userError(
        "NotImplemented",
        "teamsApp/publish requires the teams-app-publish driver. " +
          "Use the lifecycle engine to execute the publish lifecycle.",
        { source: "teamsApp/publish" }
      )
    );
  }
);

// ─── Extend to M365 ─────────────────────────────────────────

const extendToM365InputSchema = z.object({
  /** Path to the zipped app package (.zip) */
  appPackagePath: z.string().min(1),
  /** Scope for sideloading: Personal, Shared, or Tenant */
  scope: z.nativeEnum(AppScope).optional(),
});

interface ExtendToM365Result {
  titleId: string;
  appId: string;
  shareLink: string;
}

/**
 * Sideload a Teams app package to the M365 ecosystem.
 * Wraps the teamsApp/extendToM365 driver for standalone use.
 */
export const extendToM365Op = defineOperation(
  "teamsApp/extendToM365",
  extendToM365InputSchema,
  async (ctx: AtkContext, input: z.infer<typeof extendToM365InputSchema>) => {
    const packagePath = path.isAbsolute(input.appPackagePath)
      ? input.appPackagePath
      : path.resolve(ctx.projectPath ?? ".", input.appPackagePath);

    try {
      await fs.promises.access(packagePath);
    } catch {
      return err(
        userError("AppPackageNotFound", `App package not found: ${packagePath}`, {
          source: "teamsApp/extendToM365",
        })
      );
    }

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
    const scope = input.scope ?? AppScope.Personal;

    ctx.logger.info(`[teamsApp/extendToM365] Sideloading to M365 (scope: ${scope})...`);

    const sideloadRes = await service.sideLoad(packagePath, scope);
    if (sideloadRes.isErr()) return err(sideloadRes.error);

    return ok(sideloadRes.value satisfies ExtendToM365Result);
  }
);

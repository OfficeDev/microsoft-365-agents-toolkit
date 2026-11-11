// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, FxError, M365TokenProvider, ok, Result, UserError } from "@microsoft/teamsfx-api";
import isUUID from "validator/lib/isUUID";
import { AppStudioScopes } from "../../common/constants";
import { featureFlagManager, FeatureFlags } from "../../common/featureFlags";
import { getLocalizedString } from "../../common/localizeUtils";
import { metadataUtil } from "./metadataUtil";
import { pathUtils } from "./pathUtils";

export function getAppIdEnvName(output?: Record<string, string>): string | undefined {
  return output?.appId ?? output?.teamsAppId;
}

export function defaultAppIdEnvName(): string {
  return featureFlagManager.getBooleanValue(FeatureFlags.NewDeveloperPortalApis)
    ? "APP_ID"
    : "TEAMS_APP_ID";
}

export function appIdentityError(detail: string): UserError {
  return new UserError({
    source: "Teams",
    name: "InvalidAppIdentity",
    message: getLocalizedString("error.dp.invalidIdentity", detail),
  });
}

export function resolveAppResourceId(
  env: Record<string, string | undefined>,
  configuredName?: string,
  manifestFallback?: string
): Result<string | undefined, FxError> {
  const enabled = featureFlagManager.getBooleanValue(FeatureFlags.NewDeveloperPortalApis);
  const resourceId =
    (configuredName ? env[configuredName] : undefined) || env.APP_ID || env.TEAMS_APP_ID;
  const legacyResourceId = manifestFallback ?? env[configuredName ?? "TEAMS_APP_ID"];
  const separateManifestId = env.APP_MANIFEST_ID;
  if (separateManifestId !== undefined) {
    if (!resourceId || !isUUID(separateManifestId)) {
      return err(appIdentityError("APP_ID / APP_MANIFEST_ID"));
    }
    if (
      !enabled &&
      (separateManifestId !== resourceId || separateManifestId !== legacyResourceId)
    ) {
      return err(
        new UserError({
          source: "Teams",
          name: "SplitAppIdentityRequiresNewApi",
          message: getLocalizedString("error.dp.splitIdentityRequiresNewApi"),
        })
      );
    }
  }
  if (!enabled) {
    return ok(legacyResourceId);
  }
  const selected = resourceId || manifestFallback;
  if (selected && !isUUID(selected)) {
    return err(appIdentityError(configuredName ?? "APP_ID / TEAMS_APP_ID"));
  }
  return ok(selected);
}

export async function resolveProjectAppResourceId(
  projectPath: string,
  env: Record<string, string | undefined>,
  manifestFallback?: string
): Promise<Result<string | undefined, FxError>> {
  if (
    !featureFlagManager.getBooleanValue(FeatureFlags.NewDeveloperPortalApis) &&
    env.APP_MANIFEST_ID === undefined &&
    manifestFallback !== undefined
  ) {
    return ok(manifestFallback ?? env.TEAMS_APP_ID);
  }
  const yamlPath = pathUtils.getYmlFilePath(projectPath, env.TEAMSFX_ENV, true);
  let configuredName: string | undefined;
  if (yamlPath) {
    const project = await metadataUtil.parse(yamlPath);
    if (project.isErr()) return err(project.error);
    const create = project.value.provision?.driverDefs.find(
      (driver) => driver.uses === "teamsApp/create"
    );
    configuredName = getAppIdEnvName(create?.writeToEnvironmentFile);
  }
  return resolveAppResourceId(env, configuredName, manifestFallback);
}

export async function getAppTenantId(
  tokenProvider: M365TokenProvider
): Promise<Result<string, FxError>> {
  const claims = await tokenProvider.getJsonObject({ scopes: AppStudioScopes() });
  if (claims.isErr()) return err(claims.error);
  const tenantId = claims.value.tid;
  return typeof tenantId === "string" && tenantId.length > 0
    ? ok(tenantId)
    : err(appIdentityError("tid"));
}

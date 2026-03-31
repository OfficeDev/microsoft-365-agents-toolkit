// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { featureFlagManager, FeatureFlags } from "./featureFlags";

export enum SovereignCloudEnvironment {
  Public = "Public",
  GCCM = "GCC M",
  GCCH = "GCC H",
  DOD = "DoD",
}

export function getEntraEndpoint(): string {
  const sovereignCloudEnvironment = featureFlagManager.getStringValue(
    FeatureFlags.SovereignCloudEnvironment
  );
  if (
    sovereignCloudEnvironment === SovereignCloudEnvironment.GCCH ||
    sovereignCloudEnvironment === SovereignCloudEnvironment.DOD
  ) {
    return "https://login.microsoftonline.us";
  }
  return "https://login.microsoftonline.com";
}

export function getDefaultAuthorityUrl(): string {
  return `${getEntraEndpoint()}/common`;
}

export function getTenantedAuthorityUrl(tenantId: string): string {
  return `${getEntraEndpoint()}/${tenantId}`;
}

export function getSovereignCloudEnvironment(): SovereignCloudEnvironment {
  const sovereignCloudEnvironment = featureFlagManager.getStringValue(
    FeatureFlags.SovereignCloudEnvironment
  );
  if (
    sovereignCloudEnvironment &&
    Object.values(SovereignCloudEnvironment).includes(
      sovereignCloudEnvironment as SovereignCloudEnvironment
    )
  ) {
    return sovereignCloudEnvironment as SovereignCloudEnvironment;
  }
  return SovereignCloudEnvironment.Public;
}

export function isSovereignHigh(): boolean {
  const env = getSovereignCloudEnvironment();
  return env === SovereignCloudEnvironment.GCCH || env === SovereignCloudEnvironment.DOD;
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { featureFlagManager, FeatureFlags } from "./featureFlags";

enum SovereignCloudEnvironment {
  GCCM = "GCC M",
  GCCH = "GCC High",
  DOD = "DOD",
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

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export function getUsernameFromClaims(claims?: object): string {
  const tokenClaims = (claims ?? {}) as Record<string, unknown>;
  return (
    (tokenClaims.upn as string | undefined) ??
    (tokenClaims.unique_name as string | undefined) ??
    (tokenClaims.preferred_username as string | undefined) ??
    (tokenClaims.email as string | undefined) ??
    ""
  );
}

export function getInternalFlagFromTokenClaims(tokenJson: object): string {
  return getUsernameFromClaims(tokenJson).toLowerCase().endsWith("@microsoft.com")
    ? "true"
    : "false";
}

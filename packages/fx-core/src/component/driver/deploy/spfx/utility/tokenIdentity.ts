// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { M365TokenProvider } from "@microsoft/teamsfx-api";
import { GraphScopes, SPFxScopes } from "../../../../../common/constants";
import { PreparedPublicationError, PreparedPublicationJournal } from "./preparedPublication";

const tenantIdPattern =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function getClaims(value: Record<string, unknown>): Record<string, unknown> {
  const nested = value["value"];
  return isRecord(nested) ? nested : value;
}

function isGraphAudience(value: unknown): boolean {
  return (
    value === "https://graph.microsoft.com" || value === "00000003-0000-0000-c000-000000000000"
  );
}

function isSharePointAudience(value: unknown, tenantOrigin: string): boolean {
  return (
    value === "00000003-0000-0ff1-ce00-000000000000" ||
    value === tenantOrigin ||
    value === `${tenantOrigin}/`
  );
}

/**
 * Verifies that the signed-in Graph and SharePoint tokens belong to one tenant and carry the
 * expected audiences. When an expected tenant is supplied, the tokens must also belong to it.
 */
export async function verifyPreparedTokenIdentity(
  tokenProvider: M365TokenProvider,
  tenantOrigin: string,
  expectedTenantId?: string
): Promise<string> {
  const graphClaimsResult = await tokenProvider.getJsonObject({ scopes: GraphScopes });
  if (graphClaimsResult.isErr()) {
    throw new PreparedPublicationError(
      "GraphTokenIdentityUnavailable",
      "The Graph token identity could not be verified.",
      { cause: graphClaimsResult.error }
    );
  }
  const graphClaims = getClaims(graphClaimsResult.value);
  const graphTenantId = graphClaims["tid"];
  if (
    typeof graphTenantId !== "string" ||
    !tenantIdPattern.test(graphTenantId) ||
    !isGraphAudience(graphClaims["aud"])
  ) {
    throw new PreparedPublicationError(
      "GraphTokenIdentityMismatch",
      "The Graph token tenant or audience does not match the prepared publication."
    );
  }
  const sharePointClaimsResult = await tokenProvider.getJsonObject({
    scopes: SPFxScopes(tenantOrigin),
  });
  if (sharePointClaimsResult.isErr()) {
    throw new PreparedPublicationError(
      "SharePointTokenIdentityUnavailable",
      "The SharePoint token identity could not be verified.",
      { cause: sharePointClaimsResult.error }
    );
  }
  const sharePointClaims = getClaims(sharePointClaimsResult.value);
  if (
    sharePointClaims["tid"] !== graphTenantId ||
    !isSharePointAudience(sharePointClaims["aud"], tenantOrigin)
  ) {
    throw new PreparedPublicationError(
      "SharePointTokenIdentityMismatch",
      "The SharePoint token tenant or audience does not match the prepared publication."
    );
  }
  if (
    expectedTenantId !== undefined &&
    graphTenantId.toLowerCase() !== expectedTenantId.toLowerCase()
  ) {
    throw new PreparedPublicationError(
      "PreparedTenantMismatch",
      "The signed-in tenant does not match the prepared publication."
    );
  }
  return graphTenantId;
}

function normalizeCatalogUrl(value: string): string | undefined {
  try {
    const url = new URL(value);
    return `${url.origin}${url.pathname.replace(/\/+$/, "")}`.toLowerCase();
  } catch {
    return undefined;
  }
}

/**
 * Binds a prepared write to the tenant, origin and app catalog recorded during preparation.
 */
export async function verifyPreparedJournalIdentity(
  tokenProvider: M365TokenProvider,
  journal: PreparedPublicationJournal,
  tenantUrl: string,
  catalogUrl: string
): Promise<void> {
  let tenantOrigin: string;
  try {
    tenantOrigin = new URL(tenantUrl).origin;
  } catch (error) {
    throw new PreparedPublicationError(
      "InvalidPreparedTenantContext",
      "The tenant URL could not be verified.",
      { cause: error }
    );
  }
  const expectedCatalog = normalizeCatalogUrl(journal.identity.catalogUrl);
  if (
    tenantOrigin.toLowerCase() !== journal.identity.tenantOrigin.toLowerCase() ||
    expectedCatalog === undefined ||
    normalizeCatalogUrl(catalogUrl) !== expectedCatalog
  ) {
    throw new PreparedPublicationError(
      "PreparedTenantMismatch",
      "The signed-in tenant does not match the prepared publication."
    );
  }
  await verifyPreparedTokenIdentity(tokenProvider, tenantOrigin, journal.identity.tenantId);
}

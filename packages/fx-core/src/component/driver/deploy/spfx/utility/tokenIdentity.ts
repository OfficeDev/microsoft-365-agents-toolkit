// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { M365TokenProvider } from "@microsoft/teamsfx-api";
import { GraphScopes } from "../../../../../common/constants";
import { PreparedPublicationError } from "./publicationArtifacts";

const tenantIdPattern =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * Verifies the signed-in Graph token tenant and audience before using its tenant URL.
 */
export async function verifyPreparedTokenIdentity(
  tokenProvider: M365TokenProvider,
  tenantOrigin: string
): Promise<string> {
  const claimsResult = await tokenProvider.getJsonObject({ scopes: GraphScopes });
  if (claimsResult.isErr()) {
    throw new PreparedPublicationError(
      "GraphTokenIdentityUnavailable",
      "The Graph token identity could not be verified.",
      { cause: claimsResult.error }
    );
  }
  const claims = isRecord(claimsResult.value["value"])
    ? claimsResult.value["value"]
    : claimsResult.value;
  const tenantId = claims["tid"];
  const audience = claims["aud"];
  if (
    typeof tenantId !== "string" ||
    !tenantIdPattern.test(tenantId) ||
    (audience !== "https://graph.microsoft.com" &&
      audience !== "00000003-0000-0000-c000-000000000000")
  ) {
    throw new PreparedPublicationError(
      "GraphTokenIdentityMismatch",
      "The Graph token tenant or audience could not be verified."
    );
  }
  if (!tenantOrigin.startsWith("https://")) {
    throw new PreparedPublicationError(
      "InvalidPreparedTenantContext",
      "The tenant URL could not be verified."
    );
  }
  return tenantId;
}

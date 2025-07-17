// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { ConfidentialClientApplication, NodeAuthOptions } from "@azure/msal-node";
import { AppCredentialAuthConfig, OnBehalfOfCredentialAuthConfig } from "../models/configuration";
import { ClientCertificate, getAuthority } from "./utils";
import { internalLogger } from "./logger";
import { ErrorWithCode, ErrorCode } from "../core/errors";
import { createHash } from "crypto";

/**
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * @internal
 */
export function createConfidentialClientApplication(
  authentication: AppCredentialAuthConfig | OnBehalfOfCredentialAuthConfig
): ConfidentialClientApplication {
  const authority = getAuthority(authentication.authorityHost, authentication.tenantId);
  const clientCertificate: ClientCertificate | undefined = parseCertificate(
    authentication.certificateContent
  );

  const auth: NodeAuthOptions = {
    clientId: authentication.clientId,
    authority: authority,
  };

  if (clientCertificate) {
    auth.clientCertificate = clientCertificate;
  } else {
    auth.clientSecret = authentication.clientSecret;
  }

  return new ConfidentialClientApplication({
    auth,
  });
}

/**
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * @internal
 */
export function parseCertificate(
  certificateContent: string | undefined
): ClientCertificate | undefined {
  if (!certificateContent) {
    return undefined;
  }

  const certificatePattern =
    /(-+BEGIN CERTIFICATE-+)(\n\r?|\r\n?)([A-Za-z0-9+/\n\r]+=*)(\n\r?|\r\n?)(-+END CERTIFICATE-+)/;
  const match = certificatePattern.exec(certificateContent);
  if (!match) {
    const errorMsg = "The certificate content does not contain a PEM-encoded certificate.";
    internalLogger.error(errorMsg);
    throw new ErrorWithCode(errorMsg, ErrorCode.InvalidCertificate);
  }
  const thumbprint = createHash("sha256")
    .update(Buffer.from(match[3], "base64"))
    .digest("hex")
    .toUpperCase();

  return {
    thumbprintSha256: thumbprint,
    privateKey: certificateContent,
  };
}

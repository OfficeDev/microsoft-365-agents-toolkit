// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AccessToken, TokenCredential, GetTokenOptions } from "@azure/identity";
import { AppCredentialAuthConfig } from "../models/configuration";
import { internalLogger } from "../util/logger";
import { validateScopesType, formatString, getScopesArray } from "../util/utils";
import { ErrorCode, ErrorMessage, ErrorWithCode } from "../core/errors";
import { ConfidentialClientApplication } from "@azure/msal-node";
import { createConfidentialClientApplication } from "../util/utils.node";

/**
 * Represent Microsoft 365 tenant identity, and it is usually used when user is not involved like time-triggered automation job.
 *
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 *
 * @example
 * ```typescript
 * loadConfiguration(); // load configuration from environment variables
 * const credential = new AppCredential();
 * ```
 *
 * @remarks
 * Only works in in server side.
 */
export class AppCredential implements TokenCredential {
  private readonly msalClient: ConfidentialClientApplication;

  /**
   * Constructor of AppCredential.
   *
   * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
   *
   * @remarks
   * Only works in in server side.
   *
   * @param {AppCredentialAuthConfig} authConfig - The authentication configuration.
   *
   * @throws {@link ErrorCode|InvalidConfiguration} when client id, client secret or tenant id is not found in config.
   * @throws {@link ErrorCode|RuntimeNotSupported} when runtime is nodeJS.
   */
  constructor(authConfig: AppCredentialAuthConfig) {
    internalLogger.info("Create M365 tenant credential");

    const config = this.loadAndValidateConfig(authConfig);

    this.msalClient = createConfidentialClientApplication(config);
  }

  /**
   * Get access token for credential.
   *
   * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
   *
   * @example
   * ```typescript
   * await credential.getToken(["User.Read.All"]) // Get Graph access token for single scope using string array
   * await credential.getToken("User.Read.All") // Get Graph access token for single scope using string
   * await credential.getToken(["User.Read.All", "Calendars.Read"]) // Get Graph access token for multiple scopes using string array
   * await credential.getToken("User.Read.All Calendars.Read") // Get Graph access token for multiple scopes using space-separated string
   * await credential.getToken("https://graph.microsoft.com/User.Read.All") // Get Graph access token with full resource URI
   * await credential.getToken(["https://outlook.office.com/Mail.Read"]) // Get Outlook access token
   * ```
   *
   * @param {string | string[]} scopes - The list of scopes for which the token will have access.
   * @param {GetTokenOptions} options - The options used to configure any requests this TokenCredential implementation might make.
   *
   * @throws {@link ErrorCode|ServiceError} when get access token with authentication error.
   * @throws {@link ErrorCode|InternalError} when get access token with unknown error.
   * @throws {@link ErrorCode|InvalidParameter} when scopes is not a valid string or string array.
   * @throws {@link ErrorCode|RuntimeNotSupported} when runtime is nodeJS.
   *
   * @returns Access token with expected scopes.
   * Throw error if get access token failed.
   */
  async getToken(
    scopes: string | string[],
    options?: GetTokenOptions
  ): Promise<AccessToken | null> {
    let accessToken;
    validateScopesType(scopes);
    const scopesStr = typeof scopes === "string" ? scopes : scopes.join(" ");
    internalLogger.info("Get access token with scopes: " + scopesStr);

    try {
      const scopesArray = getScopesArray(scopes);
      const authenticationResult = await this.msalClient.acquireTokenByClientCredential({
        scopes: scopesArray,
      });
      if (authenticationResult) {
        accessToken = {
          token: authenticationResult.accessToken,
          expiresOnTimestamp: authenticationResult.expiresOn!.getTime(),
        };
      }
    } catch (err: any) {
      const errorMsg = "Get M365 tenant credential failed with error: " + (err.message as string);
      internalLogger.error(errorMsg);
      throw new ErrorWithCode(errorMsg, ErrorCode.ServiceError);
    }

    if (!accessToken) {
      const errorMsg = "Get M365 tenant credential access token failed with empty access token";
      internalLogger.error(errorMsg);
      throw new ErrorWithCode(errorMsg, ErrorCode.InternalError);
    }

    return accessToken;
  }

  /**
   * Load and validate authentication configuration
   *
   * @param {AuthenticationConfiguration} authConfig - The authentication configuration. Use environment variables if not provided.
   *
   * @returns Authentication configuration
   */
  private loadAndValidateConfig(config: AppCredentialAuthConfig): AppCredentialAuthConfig {
    internalLogger.verbose("Validate authentication configuration");

    if (
      config.clientId &&
      (config.clientSecret || config.certificateContent) &&
      config.tenantId &&
      config.authorityHost
    ) {
      return config;
    }

    const missingValues = [];

    if (!config.clientId) {
      missingValues.push("clientId");
    }

    if (!config.clientSecret && !config.certificateContent) {
      missingValues.push("clientSecret or certificateContent");
    }

    if (!config.tenantId) {
      missingValues.push("tenantId");
    }

    if (!config.authorityHost) {
      missingValues.push("authorityHost");
    }

    const errorMsg = formatString(
      ErrorMessage.InvalidConfiguration,
      missingValues.join(", "),
      "undefined"
    );
    internalLogger.error(errorMsg);
    throw new ErrorWithCode(errorMsg, ErrorCode.InvalidConfiguration);
  }
}

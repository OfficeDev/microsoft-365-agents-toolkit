// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Authentication configuration for TeamsUserCredential used in browser environment
 *
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 */
export type TeamsUserCredentialAuthConfig = {
  /**
   * Login page for Teams to redirect to.
   */
  initiateLoginEndpoint: string;

  /**
   * The client (application) ID of an App Registration in the tenant
   */
  clientId: string;
};

/**
 * Authentication configuration for OnBehalfOfCredential used in node environment
 *
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 */
export type OnBehalfOfCredentialAuthConfig = {
  /**
   * Hostname of AAD authority.
   */
  authorityHost: string;

  /**
   * The client (application) ID of an App Registration in the tenant
   */
  clientId: string;

  /**
   * AAD tenant id
   *
   * @readonly
   */
  tenantId: string;
} & (
  | {
      /**
       * Secret string that the application uses when requesting a token. Only used in confidential client applications. Can be created in the Azure app registration portal.
       */
      clientSecret: string;
      certificateContent?: never;
    }
  | {
      clientSecret?: never;
      /**
       * The content of a PEM-encoded public/private key certificate.
       *
       * @readonly
       */
      certificateContent: string;
    }
);

/**
 * Authentication configuration for AppCredential used in node environment
 *
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 */
export type AppCredentialAuthConfig = OnBehalfOfCredentialAuthConfig;

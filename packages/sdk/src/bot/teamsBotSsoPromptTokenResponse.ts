// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Token response provided by Teams Bot SSO prompt
 * @deprecated This package will be deprecated by 2026-07. Please use [Microsoft 365 Agents SDK](https://www.npmjs.com/package/@microsoft/agents-hosting) instead.
 */
export interface TeamsBotSsoPromptTokenResponse {
  /**
   * SSO token for user
   */
  ssoToken: string;

  /**
   * Expire time of SSO token
   */
  ssoTokenExpiration: string;
  channelId?: string;
  connectionName: string;
  expiration: string;
  properties?: {
    [propertyName: string]: any;
  };
  token: string;
}

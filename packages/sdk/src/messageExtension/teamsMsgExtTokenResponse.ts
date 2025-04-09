// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Token response provided by Teams Bot SSO prompt
 */
export interface MessageExtensionTokenResponse {
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

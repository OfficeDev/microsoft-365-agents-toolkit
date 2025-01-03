// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface OauthRegistration {
  oAuthConfigId?: string;
  /**
   * Max 128 characters
   */
  description?: string;

  clientId: string;
  clientSecret: string;

  authorizationEndpoint?: string;
  tokenExchangeEndpoint?: string;
  tokenRefreshEndpoint?: string;
  scopes: string[];

  /**
   * Teams app Id associated with the OauthRegistration, should be required if applicableToApps === "SpecificType"
   */
  m365AppId?: string;
  applicableToApps: OauthRegistrationAppType;
  /**
   * Default to be "HomeTenant"
   */
  targetAudience?: OauthRegistrationTargetAudience;
  manageableByUsers?: OauthRegistrationUser[];

  /**
   * Currently max length 1
   */
  targetUrlsShouldStartWith: string[];

  // indicating whether PKCE is enabled
  isPKCEEnabled?: boolean;

  // Identity provider, can be Custom or MicrosoftEntra
  identityProvider?: string;
}

export enum OauthRegistrationAppType {
  SpecificApp = "SpecificApp",
  AnyApp = "AnyApp",
}

export enum OauthRegistrationTargetAudience {
  HomeTenant = "HomeTenant",
  AnyTenant = "AnyTenant",
}

export interface OauthRegistrationUser {
  userId: string;
  accessType: OauthRegistrationUserAccessType;
}

export enum OauthRegistrationUserAccessType {
  Read = "Read",
  ReadWrite = "ReadWrite",
}

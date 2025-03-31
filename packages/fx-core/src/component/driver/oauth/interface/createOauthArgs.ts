// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface CreateOauthArgs {
  name: string; // The name of Api Secret
  appId: string; // Teams app id
  apiSpecPath?: string; // The location of api spec file, required if missing baseUrl, tokenUrl, or authorizationUrl
  applicableToApps?: string; // What app can access the api key. Values can be "SpecificApp" or "AnyApp". Default is "AnyApp".
  targetAudience?: string; // What tenant can access the api key. Values can be "HomeTenant" or "AnyTenant". Default is "HomeTenant".

  flow: string; // Authentication Flow. Currently only support Authorization Code Flow.
  clientId?: string; // Client id for Oauth
  clientSecret?: string; // Client secret for Oauth
  refreshUrl?: string; // Refresh url
  isPKCEEnabled?: boolean; // Whether PKCE is enabled
  identityProvider?: string; // Identity provider
  tokenExchangeMethodType?: string; // Token exchange method type

  baseUrl?: string; // The base url of the api spec. Required for both Custom Identity Provider or Microsoft Entra.
  scope?: string; // Scope(s) for Oauth, separated by a comma, optional for Custom Identity Provider, not needed for Microsoft Entra
  tokenUrl?: string; // Token url,  required for Custom Identity Provider, not needed for Microsoft Entra
  authorizationUrl?: string; // Authorization url, required for Custom Identity Provider, not needed for Microsoft Entra
}

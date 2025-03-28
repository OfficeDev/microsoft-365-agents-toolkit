// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface UpdateOauthArgs {
  name: string; // The name of Api Secret
  appId: string; // Teams app id
  apiSpecPath?: string; // The location of api spec file
  configurationId: string; // The registration id of the oauth registration
  applicableToApps?: string; // What app can access the api key. Values can be "SpecificApp" or "AnyApp". Default is "AnyApp".
  targetAudience?: string; // What tenant can access the api key. Values can be "HomeTenant" or "AnyTenant". Default is "HomeTenant".
  isPKCEEnabled?: boolean;
  clientId?: string;
  clientSecret?: string;
  identityProvider?: string;
  tokenExchangeMethodType?: string;

  refreshUrl?: string; // Refresh url
  baseUrl?: string; // The base url of the api spec. If not provided, the base url will be extracted from the api spec file.
  scope?: string; // Scope(s) for Oauth, separated by a comma
  tokenUrl?: string; // Token url
  authorizationUrl?: string; // Authorization url
}

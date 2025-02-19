// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface CreateApiKeyArgs {
  name: string; // The name of Api Secret
  appId: string; // Teams app id
  primaryClientSecret?: string; // The primary api secret
  secondaryClientSecret?: string; // The secondary api secret
  apiSpecPath: string; // The location of api spec file
  applicableToApps?: string; // What app can access the api key. Values can be "SpecificApp" or "AnyApp". Default is "AnyApp".
  targetAudience?: string; // What tenant can access the api key. Values can be "HomeTenant" or "AnyTenant". Default is "HomeTenant".
  apis?: string;
}

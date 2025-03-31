// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface UpdateApiKeyArgs {
  registrationId: string; // The registration id of the api key
  name: string; // The name of Api Secret
  appId: string; // Teams app id
  apiSpecPath?: string; // The location of api spec file
  applicableToApps?: string; // Which app can access the API key? Values can be "SpecificApp" or "AnyApp". Default is "AnyApp".
  targetAudience?: string; // Which tenant can access the API key? Values can be "HomeTenant" or "AnyTenant". Default is "AnyTenant".
  baseUrl?: string; // The base url of the api spec. If not provided, the base url will be extracted from the api spec file.
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface OauthConfigurationId {
  configurationRegistrationId: ConfigurationRegistrationId;
  resourceIdentifierUri: string;
}

export interface ConfigurationRegistrationId {
  oAuthConfigId: string;
}

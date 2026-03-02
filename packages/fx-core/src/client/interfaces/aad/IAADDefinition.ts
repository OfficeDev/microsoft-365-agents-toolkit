// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

interface Web {
  redirectUris: string[];
}

interface Spa {
  redirectUris: string[];
}

interface Oauth2PermissionScopes {
  adminConsentDescription: string;
  adminConsentDisplayName: string;
  id: string;
  isEnabled: boolean;
  type: string;
  userConsentDescription: string;
  userConsentDisplayName: string;
  value: string;
}

interface PreAuthorizedApplication {
  appId: string;
  delegatedPermissionIds: string[];
}

interface Api {
  requestedAccessTokenVersion: number;
  oauth2PermissionScopes: Oauth2PermissionScopes[];
  preAuthorizedApplications: PreAuthorizedApplication[];
}

interface AccessToken {
  name: string;
  source?: any;
  essential: boolean;
  additionalProperties: any[];
}

interface OptionalClaims {
  accessToken: AccessToken[];
}

interface ResourceAccess {
  id: string;
  type: string;
}

interface RequiredResourceAccess {
  resourceAppId?: string;
  resourceAccess?: ResourceAccess[];
}

interface PasswordCredential {
  hint?: string;
  id?: string;
  endDate?: string;
  startDate?: string;
  value?: string;
}

export interface IAADDefinition {
  displayName?: string;
  id?: string;
  appId?: string;
  identifierUris?: string[];
  web?: Web;
  signInAudience?: string;
  api?: Api;
  optionalClaims?: OptionalClaims;
  requiredResourceAccess?: RequiredResourceAccess[];
  passwordCredentials?: PasswordCredential[];
  spa?: Spa;
  serviceManagementReference?: string;
}

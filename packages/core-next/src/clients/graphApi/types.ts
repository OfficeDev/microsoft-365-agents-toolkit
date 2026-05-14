// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Subset of the AAD/Entra application definition returned by MS Graph API.
 * Only the fields we actually read/write are included.
 */
export interface AADApplication {
  id?: string; // Object ID
  appId?: string; // Client / Application ID
  displayName?: string;
  signInAudience?: string;
  identifierUris?: string[];
  web?: {
    redirectUris?: string[];
    implicitGrantSettings?: {
      enableAccessTokenIssuance?: boolean;
      enableIdTokenIssuance?: boolean;
    };
  };
  spa?: { redirectUris?: string[] };
  api?: {
    requestedAccessTokenVersion?: number;
    oauth2PermissionScopes?: OAuth2PermissionScope[];
    preAuthorizedApplications?: PreAuthorizedApplication[];
  };
  optionalClaims?: {
    accessToken?: OptionalClaim[];
    idToken?: OptionalClaim[];
  };
  requiredResourceAccess?: RequiredResourceAccess[];
  passwordCredentials?: PasswordCredential[];
  serviceManagementReference?: string;
}

export interface OAuth2PermissionScope {
  adminConsentDescription?: string;
  adminConsentDisplayName?: string;
  id: string;
  isEnabled?: boolean;
  type?: string;
  userConsentDescription?: string;
  userConsentDisplayName?: string;
  value?: string;
}

export interface PreAuthorizedApplication {
  appId: string;
  delegatedPermissionIds: string[];
}

export interface OptionalClaim {
  name: string;
  source?: string;
  essential?: boolean;
  additionalProperties?: string[];
}

export interface RequiredResourceAccess {
  resourceAppId: string;
  resourceAccess: ResourceAccess[];
}

export interface ResourceAccess {
  id: string;
  type: "Scope" | "Role";
}

export interface PasswordCredential {
  displayName?: string;
  endDateTime?: string;
  startDateTime?: string;
  secretText?: string;
  keyId?: string;
}

/**
 * Owner entry from GET /applications/{id}/owners.
 */
export interface AadOwner {
  id: string;
  displayName?: string;
  userPrincipalName?: string;
}

/**
 * Sign-in audience values accepted by Graph API.
 */
export enum SignInAudience {
  AzureADMyOrg = "AzureADMyOrg",
  AzureADMultipleOrgs = "AzureADMultipleOrgs",
  AzureADandPersonalMicrosoftAccount = "AzureADandPersonalMicrosoftAccount",
  PersonalMicrosoftAccount = "PersonalMicrosoftAccount",
}

/**
 * Bot Framework registration record used by the TDP bot API.
 */
export interface BotRegistration {
  botId?: string;
  name: string;
  description: string;
  iconUrl: string;
  messagingEndpoint: string;
  callingEndpoint: string;
  configuredChannels?: BotChannelType[];
  isSingleTenant?: boolean;
}

export enum BotChannelType {
  MicrosoftTeams = "msteams",
  M365Extensions = "m365extensions",
}

/* ------------------------------------------------------------------ */
/*  App Catalog (publishing) types                                     */
/* ------------------------------------------------------------------ */

/**
 * Publishing state of an app in the tenant app catalog.
 */
export enum PublishingState {
  submitted = "submitted",
  published = "published",
  rejected = "rejected",
}

/**
 * Published app definition from the Graph `/appCatalogs/teamsApps` endpoint.
 */
export interface PublishedAppDefinition {
  teamsAppId: string;
  displayName: string;
  publishingState: PublishingState;
  lastModifiedDateTime: Date | null;
}

/* ------------------------------------------------------------------ */
/*  Constants                                                          */
/* ------------------------------------------------------------------ */

/**
 * MS Graph API base URL (v1.0).
 * Sovereign-cloud support can be added later by making this configurable.
 */
export const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

/**
 * MS Graph API beta endpoint — required for `/appCatalogs/teamsApps` publish operations.
 */
export const GRAPH_BETA_URL = "https://graph.microsoft.com/beta";

/**
 * OAuth scopes for MS Graph application management.
 */
export function graphScopes(): string[] {
  return ["Application.ReadWrite.All"];
}

/**
 * OAuth scopes for publishing apps to the tenant app catalog.
 */
export function graphAppCatalogScopes(): string[] {
  return ["AppCatalog.ReadWrite.All"];
}

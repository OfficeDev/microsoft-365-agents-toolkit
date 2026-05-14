// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * App definition returned by the Teams Developer Portal REST API.
 * Subset of the full AppDefinition from fx-core — only fields we consume.
 */
export interface AppDefinition {
  teamsAppId?: string;
  tenantId?: string;
  appId?: string;
  appName?: string;
  shortName?: string;
  version?: string;
  createdAt?: string;
  updatedAt?: string;
}

/**
 * Staging/published app record from "GET /api/publishing/{id}".
 */
export interface PublishedAppDefinition {
  teamsAppId: string;
  displayName: string;
  publishingState: string;
  lastModifiedDateTime?: Date | null;
}

/**
 * Base URL for the Teams Developer Portal.
 * Sovereign-cloud support can be added later by making this configurable.
 */
export const TDP_BASE_URL = "https://dev.teams.microsoft.com";

/**
 * OAuth scopes required to call the Developer Portal API.
 */
export function appStudioScopes(): string[] {
  return [`${TDP_BASE_URL}/AppDefinitions.ReadWrite`];
}

/* ─── OAuth Configuration Registration ────────────────────────── */

export interface OauthRegistration {
  oAuthConfigId?: string;
  /** Max 128 characters */
  description?: string;
  clientId: string;
  clientSecret?: string;
  authorizationEndpoint?: string;
  tokenExchangeEndpoint?: string;
  tokenRefreshEndpoint?: string;
  scopes?: string[];
  m365AppId?: string;
  applicableToApps: OauthRegistrationAppType;
  targetAudience?: OauthRegistrationTargetAudience;
  targetUrlsShouldStartWith: string[];
  isPKCEEnabled?: boolean;
  identityProvider?: string;
  tokenExchangeMethodType?: TokenExchangeMethodType;
}

export enum OauthRegistrationAppType {
  SpecificApp = "SpecificApp",
  AnyApp = "AnyApp",
}

export enum OauthRegistrationTargetAudience {
  HomeTenant = "HomeTenant",
  AnyTenant = "AnyTenant",
}

export enum TokenExchangeMethodType {
  BasicAuthorizationHeader = "BasicAuthorizationHeader",
  PostRequestBody = "PostRequestBody",
}

/** Response from creating an OAuth configuration */
export interface OauthConfigurationId {
  configurationRegistrationId: { oAuthConfigId: string };
  resourceIdentifierUri: string;
}

/* ─── API Key (Secret) Registration ───────────────────────────── */

export interface ApiSecretRegistration {
  id?: string;
  /** Max 128 characters */
  description?: string;
  clientSecrets: ApiSecretRegistrationClientSecret[];
  tenantId?: string;
  targetUrlsShouldStartWith: string[];
  specificAppId?: string;
  applicableToApps: ApiSecretRegistrationAppType;
  targetAudience?: ApiSecretRegistrationTargetAudience;
  manageableByUsers?: ApiSecretRegistrationUser[];
}

export interface ApiSecretRegistrationClientSecret {
  id?: string;
  value: string;
  description?: string;
  isValueRedacted?: boolean;
  priority?: number;
}

export enum ApiSecretRegistrationAppType {
  SpecificApp = "SpecificApp",
  AnyApp = "AnyApp",
}

export enum ApiSecretRegistrationTargetAudience {
  HomeTenant = "HomeTenant",
  AnyTenant = "AnyTenant",
}

export interface ApiSecretRegistrationUser {
  userId: string;
  accessType: "Read" | "ReadWrite";
}

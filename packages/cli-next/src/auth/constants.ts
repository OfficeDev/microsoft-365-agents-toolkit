// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/** Shared MSAL client ID for both M365 and Azure interactive flows. */
export const MSAL_CLIENT_ID = "7ea7c24c-b1f6-4a20-9d11-9ae12e9e7ac0";

/** AAD endpoint base URL. */
export const AAD_ENDPOINT = "https://login.microsoftonline.com/";

/** Azure management default scope. */
export const AZURE_MANAGEMENT_SCOPE = "https://management.core.windows.net/.default";

/** Azure management scopes array. */
export const AzureScopes = (): string[] => [AZURE_MANAGEMENT_SCOPE];

/**
 * Auth service scope for M365 region endpoint discovery.
 * Equivalent to fx-core's AuthSvcScopes().
 */
export const AuthSvcScopes = (): string[] => [
  "https://authsvc.teams.microsoft.com/Region.ReadWrite",
];

/** MFA error code from AAD. */
export const MFA_CODE = "AADSTS50076";

/** MSAL token cache name for M365 account. */
export const M365_CACHE_NAME = "appStudio";

/** MSAL token cache name for Azure account. */
export const AZURE_CACHE_NAME = "azure";

/** Azure login message shown to user. */
export const azureLoginMessage = "Log in to your Azure account — opening default web browser at ";

/** M365 login message shown to user. */
export const m365LoginMessage =
  "Log in to your Microsoft 365 account — opening default web browser at ";

/** Friendly service name for keytar storage. */
export const KEYTAR_SERVICE = "Microsoft 365 Agents Toolkit";

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

export const env = {
  name: "AzureCloud",
  portalUrl: "https://portal.azure.com",
  publishingProfileUrl: "https://go.microsoft.com/fwlink/?LinkId=254432",
  managementEndpointUrl: "https://management.core.windows.net",
  resourceManagerEndpointUrl: "https://management.azure.com/",
  sqlManagementEndpointUrl: "https://management.core.windows.net:8443/",
  sqlServerHostnameSuffix: ".database.windows.net",
  galleryEndpointUrl: "https://gallery.azure.com/",
  activeDirectoryEndpointUrl: "https://login.microsoftonline.com/",
  activeDirectoryResourceId: "https://management.core.windows.net/",
  activeDirectoryGraphResourceId: "https://graph.windows.net/",
  batchResourceId: "https://batch.core.windows.net/",
  activeDirectoryGraphApiVersion: "2013-04-05",
  storageEndpointSuffix: "core.windows.net",
  keyVaultDnsSuffix: ".vault.azure.net",
  azureDataLakeStoreFileSystemEndpointSuffix: "azuredatalakestore.net",
  azureDataLakeAnalyticsCatalogAndJobEndpointSuffix: "azuredatalakeanalytics.net",
  validateAuthority: true,
};

export const unknownSubscription = "UnknownSubscription";
export const unknownSubscriptionDesc = "Cannot set subscription. Choose a correct subscription.";

export const azureLoginMessage = "Log in to your Azure account - opening default web browser at ";
export const m365LoginMessage =
  "Log in to your Microsoft 365 account - opening default web browser at ";

export const changeLoginTenantMessage = (): string =>
  `The following tenants require Multi-Factor Authentication (MFA). Use '${process.env.TEAMSFX_CLI_BIN_NAME} auth login azure --tenant TENANT_ID' to explicitly login to a tenant.`;
export const MFACode = "AADSTS50076";

export const noSubscriptionFound = "NoSubscriptionFound";
export const failToFindSubscription = "Failed to find a subscription.";
export const subscription = "subscription";
export const selectSubscription = "Select a subscription";
export const loginComponent = "login";

export const subscriptionInfoFile = "subscriptionInfo.json";
export const envDefaultJsonFile = "env.default.json";

export const sendFileTimeout = "Send success page timeout.";
export const usageError = "UsageError";
export const servicePrincipalLoginFormat = (): string =>
  `${process.env.TEAMSFX_CLI_BIN_NAME} auth login azure --service-principal --username NAME --password SECRET --tenant TENANT`;
export const codeFlowLoginFormat = (): string =>
  `${process.env.TEAMSFX_CLI_BIN_NAME} auth login azure`;

export const m365CacheName = "appStudio";

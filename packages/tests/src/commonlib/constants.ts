// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
export class TestFilePath {
  static readonly armTemplateBaseFolder = "./templates/azure";
  static readonly configFolder = "./.fx/configs";
  static readonly manifestFolder = "./templates/appPackage";

  static readonly projectSettingsFileName = "projectSettings.json";
  static readonly aadManifestTemplateFileName = "aad.template.json";
  static readonly permissionJsonFileName = "permissions.json";

  static readonly mainFileName = "main.bicep";
  static readonly provisionFileName = "provision.bicep";
  static readonly configFileName = "config.bicep";

  static readonly provisionFolder = "provision";
}

export class PluginId {
  static readonly Solution = "solution";
  static readonly Aad = "fx-resource-aad-app-for-teams";
  static readonly FrontendHosting = "fx-resource-frontend-hosting";
  static readonly SimpleAuth = "fx-resource-simple-auth";
  static readonly Bot = "fx-resource-bot";
  static readonly LocalDebug = "fx-resource-local-debug";
  static readonly AzureSQL = "fx-resource-azure-sql";
  static readonly Function = "fx-resource-function";
  static readonly Identity = "fx-resource-identity";
  static readonly Apim = "fx-resource-apim";
  static readonly KeyVault = "fx-resource-key-vault";
  static readonly AppStudio = "fx-resource-appstudio";
}

export const fileEncoding = "UTF8";

export type CliCapabilities =
  | "tab"
  | "bot"
  | "basic-message-extension"
  | "BotAndMessageExtension"
  | "TabNonSsoAndBot";
export type CliTriggerType =
  | "http-express"
  | "http-functions"
  | "http-and-timer-functions"
  | "timer-functions";

export enum Resource {
  AzureKeyVault = "azure-keyvault",
  AzureFunction = "azure-function",
  AzureApim = "azure-apim",
  AzureSql = "azure-sql",
}

export enum ResourceToDeploy {
  Spfx = "spfx",
  FrontendHosting = "frontend-hosting",
  Bot = "bot",
  Function = "function",
  Apim = "apim",
  AadManifest = "aad-manifest",
}

export enum Runtime {
  Dotnet = "dotnet",
  Node = "node",
}

export class StateConfigKey {
  // solution
  static readonly subscriptionId = "subscriptionId";
  static readonly resourceNameSuffix = "resourceNameSuffix";
  static readonly resourceGroupName = "resourceGroupName";
  static readonly teamsAppTenantId = "teamsAppTenantId";
  static readonly location = "location";

  // aad
  static readonly clientId = "clientId";
  static readonly tenantId = "tenantId";
  static readonly objectId = "objectId";
  static readonly clientSecret = "clientSecret";
  static readonly oauthAuthority = "oauthAuthority";
  static readonly oauthHost = "oauthHost";
  static readonly oauth2PermissionScopeId = "oauth2PermissionScopeId";
  static readonly applicationIdUris = "applicationIdUris";

  // app studio
  static readonly teamsAppId = "teamsAppId";

  // simple auth
  static readonly endpoint = "endpoint";
  static readonly webAppResourceId = "webAppResourceId";
  // bot
  static readonly botId = "botId";
  static readonly botPassword = "botPassword";
  static readonly botWebAppResourceId = "botWebAppResourceId";
  static readonly botResourceId = "resourceId";
  // sql
  static readonly skipAddingUser = "skipAddingUser";
  static readonly sqlEndpoint: string = "sqlEndpoint";
  static readonly databaseName: string = "databaseName";
  // function
  static readonly functionEndpoint = "functionEndpoint";
  static readonly functionAppResourceId = "functionAppResourceId";
  // frontend hosting
  static readonly domain = "domain";
  static readonly frontendResourceId = "resourceId";
  static readonly frontendEndpoint = "siteEndpoint";
  // identity
  static readonly identityClientId = "identityClientId";
  // key vault
  static readonly keyVaultResourceId = "keyVaultResourceId";

  static readonly skuName = "skuName";
}
export class ProjectSettingKey {
  static readonly solutionSettings = "solutionSettings";
  static readonly activeResourcePlugins = "activeResourcePlugins";
  static readonly capabilities = "capabilities";
}

export class provisionParametersKey {
  static readonly resourceBaseName = "resourceBaseName";
  static readonly m365ClientSecretName = "m365ClientSecretName";
  static readonly botClientSecretName = "botClientSecretName";
  static readonly simpleAuthSku = "simpleAuthSku";
}

export class EnvConstants {
  // Azure Resource
  static readonly AZURE_SUBSCRIPTION_ID = "AZURE_SUBSCRIPTION_ID";
  static readonly AZURE_RESOURCE_GROUP_NAME = "AZURE_RESOURCE_GROUP_NAME";
  static readonly RESOURCE_SUFFIX = "RESOURCE_SUFFIX";
  // Teams App
  static readonly TEAMS_APP_ID = "TEAMS_APP_ID";
  static readonly TEAMS_APP_TENANT_ID = "TEAMS_APP_TENANT_ID";
  // AAD
  static readonly AAD_APP_OBJECT_ID = "AAD_APP_OBJECT_ID";
  static readonly AAD_APP_CLIENT_ID = "AAD_APP_CLIENT_ID";
  static readonly AAD_APP_CLIENT_SECRETS = "SECRET_AAD_APP_CLIENT_SECRET";
  static readonly AAD_APP_ACCESS_AS_USER_PERMISSION_ID =
    "AAD_APP_ACCESS_AS_USER_PERMISSION_ID";
  static readonly AAD_APP_TENANT_ID = "AAD_APP_TENANT_ID";
  static readonly AAD_APP_OAUTH_AUTHORITY = "AAD_APP_OAUTH_AUTHORITY";
  static readonly AAD_APP_OAUTH_AUTHORITY_HOST = "AAD_APP_OAUTH_AUTHORITY_HOST";
  // FrontEnd
  static readonly TAB_AZURE_STORAGE_RESOURCE_ID =
    "TAB_AZURE_STORAGE_RESOURCE_ID";
  static readonly TAB_AZURE_APP_SERVICE_RESOURCE_ID =
    "TAB_AZURE_APP_SERVICE_RESOURCE_ID";
  static readonly TAB_ENDPOINT = "TAB_ENDPOINT";
  static readonly TAB_DOMAIN = "TAB_DOMAIN";
  // BOT
  static readonly BOT_ID = "BOT_ID";
  static readonly BOT_PASSWORD = "SECRET_BOT_PASSWORD";
  static readonly BOT_AZURE_APP_SERVICE_RESOURCE_ID =
    "BOT_AZURE_APP_SERVICE_RESOURCE_ID";
  static readonly BOT_AZURE_FUNCTION_RESOURCE_ID =
    "BOT_AZURE_FUNCTION_APP_RESOURCE_ID";
  static readonly BOT_DOMAIN = "BOT_DOMAIN";
  // FUNCTION
  static readonly FUNCTION_ID = "FUNCTION_RESOURCE_ID";
  static readonly FUNCTION_ID_2 = "API_FUNCTION_RESOURCE_ID";
  static readonly FUNCTION_ENDPOINT = "FUNCTION_ENDPOINT";
  static readonly FUNCTION_ENDPOINT_2 = "API_FUNCTION_ENDPOINT";
  // Container App
  static readonly AZURE_CONTAINER_APP_NAME = "AZURE_CONTAINER_APP_NAME";
  static readonly AZURE_CONTAINER_APP_RESOURCE_ID =
    "AZURE_CONTAINER_APP_RESOURCE_ID";
  // Tab
  static readonly BACKEND_APP_NAME = "BACKEND_APP_NAME";
  static readonly FRONTEND_APP_NAME = "FRONTEND_APP_NAME";

  // Agent SDK
  static readonly clientId = "clientId";
  static readonly tenantId = "tenantId";
}

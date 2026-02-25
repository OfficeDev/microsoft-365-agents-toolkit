// Main orchestration file for M365 Agent deployment
// This deploys: Managed Identity → App Service → Azure Bot → App Registration

targetScope = 'resourceGroup'

@maxLength(20)
@minLength(4)
@description('Used to generate names for all resources')
param resourceBaseName string

@maxLength(42)
@description('Display name for the bot')
param botDisplayName string

@description('Azure AI Foundry Agent ID')
param agentId string

@description('Azure AI Foundry Project Endpoint')
param azureAIFoundryProjectEndpoint string

@description('Location for all resources')
param location string = resourceGroup().location

@description('The SKU for the App Service Plan')
@allowed([
  'F1'
  'B1'
  'B2'
  'B3'
  'S1'
  'S2'
  'S3'
  'P1v2'
  'P2v2'
  'P3v2'
])
param webAppSKU string = 'B1'

@description('The SKU for the Bot Service')
@allowed([
  'F0'
  'S1'
])
param botServiceSku string = 'F0'

@description('Tenant ID for the Entra ID application')
param tenantId string = tenant().tenantId

@description('Enable Application Insights')
param enableAppInsights bool = true

@description('Additional app settings for the Web App')
param additionalAppSettings array = []

// Generate resource names
var identityName = '${resourceBaseName}-identity'
var webAppName = '${resourceBaseName}-app'
var botServiceName = '${resourceBaseName}-bot'
var aadAppName = '${resourceBaseName}-UserAuth'

// Setp 0: GUID ENCODING: Encode Tenant ID 
module guidEncoder 'modules/guid-encoder.bicep' =  {
  name: 'encode-tenant-guid-local'
  params: {
    guidToEncode: tenantId
    location: location
  }
}


// Step 1: Create User Assigned Managed Identity for the bot
module botIdentity 'modules/bot-managedidentity.bicep' = {
  name: 'deploy-bot-identity'
  params: {
    identityName: identityName
    location: location
  }
}

// Step 1.5: Create Application Insights with Managed Identity (if enabled)
module appInsights 'modules/appinsights.bicep' = if (enableAppInsights) {
  name: 'deploy-app-insights'
  params: {
    resourceBaseName: resourceBaseName
    location: location
    identityPrincipalId: botIdentity.outputs.identityPrincipalId
    applicationType: 'web'
  }
}

// Step 2: Create App Service with the managed identity
module appService 'modules/appservice.bicep' = {
  name: 'deploy-app-service'
  params: {
    resourceBaseName: resourceBaseName
    location: location
    serverfarmsName: '${resourceBaseName}-plan'
    webAppName: webAppName
    webAppSKU: webAppSKU
    MSIid: botIdentity.outputs.identityId
    enableAppInsights: enableAppInsights
    appInsightsConnectionString: appInsights.?outputs.?appInsightsConnectionString ?? ''
    // Bot Configuration (for appsettings.json template variables)
    botId: botIdentity.outputs.identityClientId
    botTenantId: tenantId
    oauthConnectionName: 'aifoundryaccess'
    // AI Services Configuration (optional - add to parameters if needed)
    azureAIFoundryEndpoint: azureAIFoundryProjectEndpoint
    azureAIAgentId: agentId
    additionalAppSettings: additionalAppSettings
  }
}

// Step 3: Create Azure Bot Service
module azureBot 'modules/azurebot.bicep' = {
  name: 'deploy-azure-bot'
  params: {
    resourceBaseName: resourceBaseName
    botDisplayName: botDisplayName
    botServiceName: botServiceName
    botServiceSku: botServiceSku
    identityResourceId: botIdentity.outputs.identityId
    identityClientId: botIdentity.outputs.identityClientId
    identityTenantId: tenantId
    botAppDomain: appService.outputs.webAppHostName
  }
}

// Step 4: Create App Registration with all required parameters
module appRegistration 'modules/app-registration.bicep' = {
  name: 'deploy-app-registration'
  params: {
    aadAppName: aadAppName
    botId: botIdentity.outputs.identityClientId
    tenantId: tenantId
    encodedTenantId: guidEncoder.outputs.encodedGuid
  }
  dependsOn: [
    azureBot
  ]
}

// Step 5: Configure OAuth Connection with Azure AD v2 and Federated Credentials
module botOAuthConnection 'modules/bot-oauth-connection.bicep' = {
  name: 'deploy-bot-oauth-connection'
  params: {
    botServiceName: botServiceName
    connectionName: 'SsoConnection'
    aadAppId: appRegistration.outputs.aadAppId
    aadAppIdUri: appRegistration.outputs.aadAppIdUri
    federatedCredentialName: appRegistration.outputs.fciName
    scopes: '${appRegistration.outputs.aadAppIdUri}/access_as_user'
    tenantId: tenantId
    location: 'global'
  }
}

// Step 6: Configure OAuth Connection with Azure AD v2 and Federated Credentials For Azure AI Foundy
// Note: This is used to ask directly the access token to ABS vs hosting a secure client and do an OBO Flow. It is simpler, leaner and ABS handle concent & caching for us.
module botOAuthConnectionAIFoundry 'modules/bot-oauth-connection.bicep' = {
  name: 'deploy-bot-oauth-connection-aifoundry'
  params: {
    botServiceName: botServiceName
    connectionName: 'aifoundryaccess'
    aadAppId: appRegistration.outputs.aadAppId
    aadAppIdUri: appRegistration.outputs.aadAppIdUri
    federatedCredentialName: appRegistration.outputs.fciName
    scopes: 'https://ai.azure.com/user_impersonation'
    tenantId: tenantId
    location: 'global'
  }
}


// ========================================
// STEP 7: Create Service Principal for SSO App (First-time only)
// ========================================
// The SSO App is created by M365 Agents Toolkit with a client secret
// We create its service principal after SSO app registration to avoid replication timing issues
module SSOServicePrincipal 'modules/service-principal.bicep' = {
  name: 'deploy-sso-service-principal-local'
  params: {
    appId: appRegistration.outputs.aadAppId
  }
}


// Outputs for reference and further configuration
output resourceBaseName string = resourceBaseName
output location string = location

// Identity outputs
output identityName string = identityName
output identityId string = botIdentity.outputs.identityClientId
output identityPrincipalId string = botIdentity.outputs.identityPrincipalId

// App Service outputs
output webAppName string = appService.outputs.webAppName
output webAppId string = appService.outputs.webAppId
output BOT_DOMAIN string = appService.outputs.webAppHostName
output webAppUrl string = 'https://${appService.outputs.webAppHostName}'
output appServicePlanId string = appService.outputs.appServicePlanId

// Bot Service outputs
output BOT_ID string = botIdentity.outputs.identityClientId
output botServiceName string = botServiceName
output bot_Endpoint string = 'https://${appService.outputs.webAppHostName}/api/messages'
output Oauth_Connection_Name string =botOAuthConnection.name
output AIFoundry_Connection_Name string = botOAuthConnectionAIFoundry.name

// App Registration outputs
output SSO_APP_ID string = appRegistration.outputs.aadAppId
output SSO_APP_ID_URI string = appRegistration.outputs.aadAppIdUri
output federatedCredentialName string = appRegistration.outputs.fciName
// Note: fciSubject is used internally for OAuth connection but not exposed as output

// Application Insights outputs (if enabled)
output appInsightsName string = appInsights.?outputs.?appInsightsName ?? ''
output appInsightsConnectionString string = appInsights.?outputs.?appInsightsConnectionString ?? ''
output appInsightsInstrumentationKey string = appInsights.?outputs.?appInsightsInstrumentationKey ?? ''
output logAnalyticsWorkspaceName string = appInsights.?outputs.?logAnalyticsWorkspaceName ?? ''




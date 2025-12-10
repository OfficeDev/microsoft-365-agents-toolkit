@maxLength(20)
@minLength(4)
@description('Used to generate names for all resources in this file')
param resourceBaseName string

@maxLength(42)
param botDisplayName string

param botServiceName string = resourceBaseName
param botServiceSku string = 'F0'
param identityResourceId string
param identityClientId string
param identityTenantId string
param botAppDomain string

param deployAppService bool
param botConnectionClientId string
@secure()
param botConnectionClientSecret string
param botConnectionTenantId string

// Register your web service as a bot with the Bot Framework
resource botService 'Microsoft.BotService/botServices@2021-03-01' = {
  kind: 'azurebot'
  location: 'global'
  name: botServiceName
  properties: {
    displayName: botDisplayName
    endpoint: 'https://${botAppDomain}/api/messages'
    msaAppId: deployAppService ? identityClientId : botConnectionClientId
    msaAppMSIResourceId: deployAppService ? identityResourceId : ''
    msaAppTenantId: identityTenantId
    msaAppType: deployAppService ? 'UserAssignedMSI' : 'SingleTenant'
  }
  sku: {
    name: botServiceSku
  }
}

// Connect the bot service to Microsoft Teams
resource botServiceMsTeamsChannel 'Microsoft.BotService/botServices/channels@2021-03-01' = {
  parent: botService
  location: 'global'
  name: 'MsTeamsChannel'
  properties: {
    channelName: 'MsTeamsChannel'
  }
}

// Add OAuth Connection settings for the botService
resource botServiceOAuthConnection 'Microsoft.BotService/botServices/Connections@2021-03-01' = {
  parent: botService
  location: 'global'
  name: 'GraphConnection'
  properties: {
    clientId: botConnectionClientId
    clientSecret: botConnectionClientSecret
    scopes: 'Files.Read.All Sites.Read.All'
    serviceProviderDisplayName: 'Azure Active Directory v2'
    serviceProviderId: '30dd229c-58e3-4a48-bdfd-91ec48eb906c'
    parameters: [
      {
        key: 'tenantId'
        value: botConnectionTenantId
      }
    ]
  }
}

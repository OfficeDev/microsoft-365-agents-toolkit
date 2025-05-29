@maxLength(20)
@minLength(4)
@description('Used to generate names for all resources in this file')
param resourceBaseName string

@description('Required when create Azure Bot service')
param botAadAppClientId string

@secure()
@description('Required by Bot Framework package in your bot project')
param botAadAppClientSecret string

@maxLength(42)
param botDisplayName string

param botAppDomain string
param oauthConnectionName string

module azureBotRegistration './botRegistration/azurebot.bicep' = {
  name: 'Azure-Bot-registration'
  params: {
    resourceBaseName: resourceBaseName
    botAadAppClientId: botAadAppClientId
    botAppDomain: botAppDomain
    botDisplayName: botDisplayName
    botAddAppClientSecret: botAadAppClientSecret
    oauthConnectionName: oauthConnectionName
  }
}

@maxLength(20)
@minLength(4)
@description('Used to generate names for all resources in this file')
param resourceBaseName string

param tabAadAppClientId string
param tabAadAppOauthAuthorityHost string
param tabAadAppTenantId string
@secure()
param tabAadAppClientSecret string

param webAppSKU string

param serverfarmsName string = resourceBaseName
param webAppName string = resourceBaseName
param location string = resourceGroup().location

// Compute resources for your Web App
resource serverfarm 'Microsoft.Web/serverfarms@2021-02-01' = {
  kind: 'app'
  location: location
  name: serverfarmsName
  sku: {
    name: webAppSKU
  }
}

// Web App that hosts your bot
resource webApp 'Microsoft.Web/sites@2021-02-01' = {
  kind: 'app'
  location: location
  name: webAppName
  properties: {
    serverFarmId: serverfarm.id
    httpsOnly: true
    siteConfig: {
      ftpsState: 'FtpsOnly'
    }
  }
}

resource  webAppConfig  'Microsoft.Web/sites/config@2021-02-01' = {
  name: '${webAppName}/appsettings'
  properties: {
    WEBSITE_RUN_FROM_PACKAGE: '1'
    TeamsFx__Authentication__ClientId: tabAadAppClientId
    TeamsFx__Authentication__ClientSecret: tabAadAppClientSecret
    TeamsFx__Authentication__InitiateLoginEndpoint: 'https://${webApp.properties.defaultHostName}/auth-start.html'
    TeamsFx__Authentication__OAuthAuthority: uri(tabAadAppOauthAuthorityHost, tabAadAppTenantId)
  }
}

// The output will be persisted in .env.{envName}. Visit https://aka.ms/teamsfx-actions/arm-deploy for more details.
output TAB_AZURE_APP_SERVICE_RESOURCE_ID  string = webApp.id
output TAB_DOMAIN  string = webApp.properties.defaultHostName
output TAB_HOSTNAME  string = webApp.properties.defaultHostName
output TAB_ENDPOINT  string = 'https://${webApp.properties.defaultHostName}'

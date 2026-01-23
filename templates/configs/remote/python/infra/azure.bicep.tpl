@maxLength(20)
@minLength(4)
@description('Used to generate names for all resources in this file')
param resourceBaseName string

param webAppSku string
param linuxFxVersion string

param serverfarmsName string = resourceBaseName
param webAppName string = resourceBaseName
param location string = resourceGroup().location
param pythonVersion string = linuxFxVersion
{{#hasAzureBot}}
@maxLength(42)
param botDisplayName string
{{/hasAzureBot}}

{{#hasAzureBot}}
resource identity 'Microsoft.ManagedIdentity/userAssignedIdentities@2023-01-31' = {
  location: location
  name: resourceBaseName
}
{{/hasAzureBot}}


resource serverfarm 'Microsoft.Web/serverfarms@2021-02-01' = {
  kind: 'app,linux'
  location: location
  name: serverfarmsName
  sku: {
    name: webAppSku
  }
  properties:{
    reserved: true
  }
}

resource webApp 'Microsoft.Web/sites@2021-02-01' = {
  kind: 'app,linux'
  location: location
  name: webAppName
  properties: {
    serverFarmId: serverfarm.id
    httpsOnly: true
    siteConfig: {
      alwaysOn: true
      appCommandLine: 'python app.py'
      linuxFxVersion: linuxFxVersion
      appSettings: [
        {
          name: 'WEBSITES_CONTAINER_START_TIME_LIMIT'
          value: '900'
        }
        {
          name: 'SCM_DO_BUILD_DURING_DEPLOYMENT'
          value: 'true'
        }
{{#hasAzureBot}}
        {
          name: 'CLIENT_ID'
          value: identity.properties.clientId
        }
        {
          name: 'TENANT_ID'
          value: identity.properties.tenantId
        }
        {
          name: 'BOT_TYPE'
          value: 'UserAssignedMsi'
        }
{{/hasAzureBot}}
      ]
      ftpsState: 'FtpsOnly'
    }
  }
{{#hasAzureBot}}
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${identity.id}': {}
    }
  }
{{/hasAzureBot}}
}

{{#hasAzureBot}}
// Register your web service as a bot with the Bot Framework
module azureBotRegistration './botRegistration/azurebot.bicep' = {
  name: 'Azure-Bot-registration'
  params: {
    resourceBaseName: resourceBaseName
    botDisplayName: botDisplayName
    identityResourceId: identity.id
    identityClientId: identity.properties.clientId
    identityTenantId: identity.properties.tenantId
    botAppDomain: webApp.properties.defaultHostName
  }
}
{{/hasAzureBot}}

// The output will be persisted in .env.{envName}. Visit https://aka.ms/teamsfx-actions/arm-deploy for more details.
output AZURE_APP_SERVICE_RESOURCE_ID string = webApp.id // used in deploy stage
output BOT_DOMAIN string = webApp.properties.defaultHostName
output BOT_ENDPOINT string = 'https://${webApp.properties.defaultHostName}'
{{#hasAzureBot}}
output BOT_ID string = identity.properties.clientId
output BOT_TENANT_ID string = identity.properties.tenantId
{{/hasAzureBot}}

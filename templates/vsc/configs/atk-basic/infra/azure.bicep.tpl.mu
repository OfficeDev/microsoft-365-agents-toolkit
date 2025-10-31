@maxLength(20)
@minLength(4)
@description('Used to generate names for all resources in this file')
param resourceBaseName string

param webAppSku string

param serverfarmsName string = resourceBaseName
param webAppName string = resourceBaseName
param location string = resourceGroup().location

<<#bot>>
@maxLength(42)
param botDisplayName string

{{#useOpenAI}}
@secure()
param openAIKey string
{{/useOpenAI}}
{{#useAzureOpenAI}}
@secure()
param azureOpenAIKey string

@secure()
param azureOpenAIEndpoint string

@secure()
param azureOpenAIDeploymentName string

@secure()
param azureOpenAIEmbeddingDeploymentName string
{{/useAzureOpenAI}}

resource identity 'Microsoft.ManagedIdentity/userAssignedIdentities@2023-01-31' = {
  location: location
  name: identityName
}
<</bot>>

resource serverfarm 'Microsoft.Web/serverfarms@2021-02-01' = {
  kind: 'app'
  location: location
  name: serverfarmsName
  sku: {
    name: webAppSku
  }
}

resource webApp 'Microsoft.Web/sites@2021-02-01' = {
  kind: 'app'
  location: location
  name: webAppName
  properties: {
    serverFarmId: serverfarm.id
    httpsOnly: true
<<#bot>>
    alwaysOn: true
<</bot>>
    siteConfig: {
      appSettings: [
        {
          name: 'WEBSITE_RUN_FROM_PACKAGE'
          value: '1' // Run Azure App Service from a package file
        }
        {
          name: 'WEBSITE_NODE_DEFAULT_VERSION'
          value: '~22' // Set NodeJS version to 22.x for your site
        }
        {
          name: 'RUNNING_ON_AZURE'
          value: '1'
        }
<<#bot>>
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
        {{#useOpenAI}}
        {
          name: 'OPENAI_API_KEY'
          value: openAIKey
        }
        {{/useOpenAI}}
        {{#useAzureOpenAI}}
        {
          name: 'AZURE_OPENAI_API_KEY'
          value: azureOpenAIKey
        }
        {
          name: 'AZURE_OPENAI_ENDPOINT'
          value: azureOpenAIEndpoint
        }
        {
          name: 'AZURE_OPENAI_DEPLOYMENT_NAME'
          value: azureOpenAIDeploymentName
        }
        {
          name: 'AZURE_OPENAI_EMBEDDING_DEPLOYMENT_NAME'
          value: azureOpenAIEmbeddingDeploymentName
        }
        {{/useAzureOpenAI}}
<</bot>>
      ]
      ftpsState: 'FtpsOnly'
    }
  }
<<#bot>>
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${identity.id}': {}
    }
  }
<</bot>>
}

<<#bot>>
// Register your web service as a bot with the Bot Framework
module azureBotRegistration './botRegistration/azurebot.bicep' = {
  name: 'Azure-Bot-registration'
  params: {
    resourceBaseName: resourceBaseName
    identityClientId: identity.properties.clientId
    identityResourceId: identity.id
    identityTenantId: identity.properties.tenantId
    botAppDomain: webApp.properties.defaultHostName
    botDisplayName: botDisplayName
  }
}
<</bot>>

// The output will be persisted in .env.{envName}. Visit https://aka.ms/teamsfx-actions/arm-deploy for more details.
output AZURE_APP_SERVICE_RESOURCE_ID string = webApp.id // used in deploy stage
output APP_DOMAIN string = webApp.properties.defaultHostName
output APP_ENDPOINT string = 'https://${webApp.properties.defaultHostName}'
<<#bot>>
output BOT_ID string = identity.properties.clientId
output BOT_TENANT_ID string = identity.properties.tenantId
<</bot>>

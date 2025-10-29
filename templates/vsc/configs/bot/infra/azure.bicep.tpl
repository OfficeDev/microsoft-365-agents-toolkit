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

resource webApp 'Microsoft.Web/sites@2021-02-01' = {
  properties: {
    siteConfig: {
      alwaysOn: true
      appSettings: [
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
      ]
    }
  }
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${identity.id}': {}
    }
  }
}

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

output BOT_ID string = identity.properties.clientId
output BOT_TENANT_ID string = identity.properties.tenantId
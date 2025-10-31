{
    "$schema": "https://schema.management.azure.com/schemas/2015-01-01/deploymentParameters.json#",
    "contentVersion": "1.0.0.0",
    "parameters": {
      "resourceBaseName": {
        "value": "app${{RESOURCE_SUFFIX}}"
      },
      "webAppSku": {
        "value": "B1"
      }
<<#bot>>
      "botDisplayName": {
        "value": "{{appName}}"
      },
      {{#useOpenAI}}
      "openAIKey": {
        "value": "${{SECRET_OPENAI_API_KEY}}"
      },
      {{/useOpenAI}}
      {{#useAzureOpenAI}}
      "azureOpenAIKey": {
        "value": "${{SECRET_AZURE_OPENAI_API_KEY}}"
      },
      "azureOpenAIEndpoint": {
        "value": "${{AZURE_OPENAI_ENDPOINT}}"
      },
      "azureOpenAIDeploymentName": {
        "value": "${{AZURE_OPENAI_DEPLOYMENT_NAME}}"
      },
      "azureOpenAIEmbeddingDeploymentName": {
        "value": "${{AZURE_OPENAI_EMBEDDING_DEPLOYMENT_NAME}}"
      },
      {{/useAzureOpenAI}}
<</bot>>
    }
  }
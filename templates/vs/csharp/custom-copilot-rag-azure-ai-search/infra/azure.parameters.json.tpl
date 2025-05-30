{
    "$schema": "https://schema.management.azure.com/schemas/2015-01-01/deploymentParameters.json#",
    "contentVersion": "1.0.0.0",
    "parameters": {
      "resourceBaseName": {
        "value": "bot${{RESOURCE_SUFFIX}}"
      },
{{#useOpenAI}}
      "openAIApiKey": {
        "value": "${{SECRET_OPENAI_API_KEY}}"
      },
      "openAIEmbeddingModel": {
        "value": "${{OPENAI_EMBEDDING_MODEL}}"
      },
{{/useOpenAI}}
{{#useAzureOpenAI}}
      "azureOpenAIApiKey": {
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
      "AISearchApiKey": {
        "value": "${{SECRET_AI_SEARCH_API_KEY}}"
      },
      "AISearchEndpoint": {
        "value": "${{AI_SEARCH_ENDPOINT}}"
      },
      "webAppSKU": {
        "value": "B1"
      },
      "botDisplayName": {
        "value": "{{appName}}"
      }
    }
  }
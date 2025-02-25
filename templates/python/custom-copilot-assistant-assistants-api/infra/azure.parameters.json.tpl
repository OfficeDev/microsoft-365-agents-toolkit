{
  "$schema": "https://schema.management.azure.com/schemas/2015-01-01/deploymentParameters.json#",
  "contentVersion": "1.0.0.0",
  "parameters": {
    "resourceBaseName": {
      "value": "bot${{RESOURCE_SUFFIX}}"
    },
    {{#useOpenAI}}
    "openaiKey": {
      "value": "${{SECRET_OPENAI_API_KEY}}"
    },
    "assistantId": {
      "value": "${{OPENAI_ASSISTANT_ID}}"
    },
    {{/useOpenAI}}
    {{#useAzureOpenAI}}
    "azureOpenaiKey": {
      "value": "${{SECRET_AZURE_OPENAI_API_KEY}}"
    },
    "azureOpenaiModelDeploymentName" : {
      "value": "${{AZURE_OPENAI_MODEL_DEPLOYMENT_NAME}}"
    },
    "azureOpenaiEndpoint" : {
      "value": "${{AZURE_OPENAI_ENDPOINT}}"
    },
    "assistantId": {
      "value": "${{AZURE_OPENAI_ASSISTANT_ID}}"
    },
    {{/useAzureOpenAI}}
    "webAppSKU": {
      "value": "B1"
    },
    "botDisplayName": {
      "value": "{{appName}}"
    },
    "linuxFxVersion": {
      "value": "PYTHON|3.11"
    }
  }
}
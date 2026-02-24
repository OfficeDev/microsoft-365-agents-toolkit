{
    "$schema": "https://schema.management.azure.com/schemas/2015-01-01/deploymentParameters.json#",
    "contentVersion": "1.0.0.0",
    "parameters": {
      "resourceBaseName": {
        "value": "bot${{RESOURCE_SUFFIX}}"
      },
      "botDisplayName": {
        "value": "{{appName}}${{APP_NAME_SUFFIX}}"
      },
      "webAppSKU": {
        "value": "B1"
      },
      "botServiceSku": {
        "value": "F0"
      },
      "enableAppInsights": {
        "value": true
      },
      "agentId": {
        "value": "${{AGENT_ID}}"
      },
      "azureAIFoundryProjectEndpoint": {
        "value": "${{AZURE_AI_FOUNDRY_PROJECT_ENDPOINT}}"
      },
      "ssoAppId": {
        "value": "${{SSO_APP_ID}}"
      },
      "ssoUniqueName": {
        "value": "{{appName}}-${{RESOURCE_SUFFIX}}-${{APP_NAME_SUFFIX}}-UserAuth"
      }
    }
  }
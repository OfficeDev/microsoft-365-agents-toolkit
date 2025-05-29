{
  "$schema": "https://schema.management.azure.com/schemas/2015-01-01/deploymentParameters.json#",
  "contentVersion": "1.0.0.0",
  "parameters": {
    "resourceBaseName": {
      "value": "bot-${{RESOURCE_SUFFIX}}-${{TEAMSFX_ENV}}"
    },
    "botAadAppClientId": {
      "value": "${{BOT_ID}}"
    },
    "botAadAppClientSecret": {
      "value": "${{SECRET_BOT_PASSWORD}}"
    },
    "botDisplayName": {
      "value": "{{appName}}-${{TEAMSFX_ENV}}"
    },
    "botAppDomain": {
      "value": "${{BOT_DOMAIN}}"
    },
    "oauthConnectionName": {
      "value": "${{OAUTH_CONNECTION_NAME}}"
    }
  }
}

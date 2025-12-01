{
  "$schema": "https://schema.management.azure.com/schemas/2015-01-01/deploymentParameters.json#",
  "contentVersion": "1.0.0.0",
  "parameters": {
    "resourceBaseName": {
      "value": "plugin${{RESOURCE_SUFFIX}}"
    },
    "functionAppSKU": {
      "value": "Y1"
    },
    "aadAppClientId": {
      "value": "${{APP_CLIENT_ID}}"
    },
    "aadAppTenantId": {
      "value": "${{APP_TENANT_ID}}"
    },
    "aadAppOauthAuthorityHost": {
      "value": "${{APP_AUTHORITY_HOST}}"
    }
  }
}
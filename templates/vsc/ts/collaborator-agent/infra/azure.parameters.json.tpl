{
  "$schema": "https://schema.management.azure.com/schemas/2015-01-01/deploymentParameters.json#",
  "contentVersion": "1.0.0.0",
  "parameters": {
    "resourceBaseName": {
      "value": "bot${{RESOURCE_SUFFIX}}"
    },
    "webAppSKU": {
      "value": "B1"
    },
    "botDisplayName": {
      "value": "{{appName}}"
    },
    "AOAI_ENDPOINT": {
      "value": "${{AOAI_ENDPOINT}}"
    },
    "AOAI_API_KEY": {
      "value": "${{AOAI_API_KEY}}"
    },
    "AOAI_MODEL": {
      "value": "${{AOAI_MODEL}}"
    },
    "sqlAdminPassword": {
      "value": "${{SQL_ADMIN_PASSWORD}}"
    }
  }
}
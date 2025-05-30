{
    "$schema": "https://schema.management.azure.com/schemas/2015-01-01/deploymentParameters.json#",
    "contentVersion": "1.0.0.0",
    "parameters": {
      "resourceBaseName": {
        "value": "notification${{RESOURCE_SUFFIX}}"
      },
      "functionAppSKU": {
        "value": "B1"
      },
      "botDisplayName": {
        "value": "{{appName}}"
      }
    }
  }
{
  "$schema": "https://schema.management.azure.com/schemas/2015-01-01/deploymentParameters.json#",
  "contentVersion": "1.0.0.0",
  "parameters": {
{{#hasAzureBot}}
    "botDisplayName": {
      "value": "{{appName}}"
    },
{{/hasAzureBot}}
    "resourceBaseName": {
      "value": "app$\{{RESOURCE_SUFFIX}}"
    },
    "webAppSku": {
      "value": "B3"
    },
    "linuxFxVersion": {
      "value": "PYTHON|3.12"
    }
  }
}
{
  "$schema": "https://schema.management.azure.com/schemas/2019-04-01/deploymentParameters.json#",
  "contentVersion": "1.0.0.0",
  "parameters": {
    "botName": {
      "value": "{{appName}}-${{RESOURCE_SUFFIX}}-${{APP_NAME_SUFFIX}}"
    },
    "botId": {
      "value": "${{BOT_ID}}"
    },
    "botEndpoint": {
      "value": "${{BOT_ENDPOINT}}"
    },
    "tenantId": {
      "value": "${{TEAMS_APP_TENANT_ID}}"
    },
    "botServiceNameState": {
      "value": "${{BOT_SERVICE_NAME}}"
    },
    "ssoAppId": {
      "value": "${{SSO_APP_ID}}"
    },
    "ssoAppObjectId": {
      "value": "${{SSO_APP_OBJECT_ID}}"
    },
    "ssoAppName": {
      "value": "{{appName}}-${{RESOURCE_SUFFIX}}-${{APP_NAME_SUFFIX}}-UserAuth"
    }
  }
}

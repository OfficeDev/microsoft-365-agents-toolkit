{
  "id": "${{AAD_APP_OBJECT_ID}}",
  "appId": "${{AAD_APP_CLIENT_ID}}",
  "name": "{{appName}}-${{TEAMSFX_ENV}}",
  "accessTokenAcceptedVersion": 2,
  "signInAudience": "AzureADMyOrg",
  "requiredResourceAccess": [
    {
      "resourceAppId": "Microsoft Graph",
      "resourceAccess": [
        {
          "id": "ExternalConnection.ReadWrite.OwnedBy",
          "type": "Role"
        },
        {
          "id": "ExternalItem.ReadWrite.OwnedBy",
          "type": "Role"
        }
      ]
    }
  ]
}

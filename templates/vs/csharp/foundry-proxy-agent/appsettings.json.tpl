{
  "AgentApplication": {
    "StartTypingTimer": true,
    "RemoveRecipientMention": false,
    "NormalizeMentions": false,
    "UserAuthorization": {
      "AutoSignIn": false,
      "Handlers": {
        "SSO": {
          "Settings": {
            "AzureBotOAuthConnectionName": "{{OAUTHCONNECTIONNAME}}"
          }
        }
      }
    }
  },

  "TokenValidation": {
    "Audiences": [
      "{{BOT_ID}}"
    ]
  },

  "Connections": {
    "ServiceConnection": {
      "Settings": {
        "AuthType": "UserManagedIdentity",
        "ClientId": "{{BOT_ID}}",
        "TenantId": "{{BOT_TENANT_ID}}",
        "Scopes": [
          "https://api.botframework.com/.default"
        ]
      }
    }
  },
  "ConnectionsMap": [
    {
      "ServiceUrl": "*",
      "Connection": "BotServiceConnection"
    }
  ],

  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft.AspNetCore": "Warning",
      "Microsoft.Agents": "Warning",
      "Microsoft.Hosting.Lifetime": "Information"
    }
  },

  "AIServices": {
    "AzureAIFoundryProjectEndpoint": "",
    "AgentID": ""
  }
}

{
  "AgentApplication": {
    "StartTypingTimer": true,
    "RemoveRecipientMention": false,
    "NormalizeMentions": false,
    "UserAuthorization": {
      "Default": "graph",
      "AutoSignIn": false,
      "Handlers": {
        "graph": {
          "Settings": {
            "AzureBotOAuthConnectionName": "GraphConnection" //name of the connection in Azure Bot Service,
          }
        }
      }
    }
  },

  "TokenValidation": {
    "Audiences": [
      "{{BOT_ID}}" // this is the Client ID used for the Azure Bot
    ]
  },

  "Connections": {
    "BotServiceConnection": {
      "Assembly": "Microsoft.Agents.Authentication.Msal",
      "Type": "MsalAuth",
      "Settings": {
        "AuthType": "UserManagedIdentity", // this is the AuthType for the connection, valid values can be found in Microsoft.Agents.Authentication.Msal.Model.AuthTypes.
        "ClientId": "{{BOT_ID}}", // this is the Client ID used for the connection.
        "TenantId": "{{BOT_TENANT_ID}}",
        "Scopes": ["https://api.botframework.com/.default"]
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

  // This is the configuration for the AI services, use environeent variables or user secrets to store sensitive information.
  // Do not store sensitive information in this file
  "Azure": {
    "OpenAIApiKey": "",
    "OpenAIEndpoint": "",
    "OpenAIDeploymentName": ""
  }
}

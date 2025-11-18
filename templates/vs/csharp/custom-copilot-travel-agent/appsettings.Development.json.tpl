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
            "AzureBotOAuthConnectionName": "GraphConnection"
          }
        }
      }
    }
  },

  "TokenValidation": {
    "Audiences": {
      "ClientId": ""
    }
  },

  "Connections": {
    "BotServiceConnection": {
      "Settings": {
        "AuthType": "ClientSecret",
        "ClientId": "",
        "ClientSecret": "",
        "AuthorityEndpoint": "",
        "Scopes": [
            "https://api.botframework.com/.default"
        ],
		"TenantId": ""
      }
    }
  },
  "ConnectionsMap": [
    {
      "ServiceUrl": "*",
      "Connection": "BotServiceConnection"
    }
  ],
  {{#useAzureOpenAI}}
  "Azure": {
    "OpenAIApiKey": "",
    "OpenAIEndpoint": "",
    "OpenAIDeploymentName": ""
  }
  {{/useAzureOpenAI}}
  {{#useOpenAI}}
  "OpenAI": {
    "ApiKey": ""
  }
  {{/useOpenAI}}
}

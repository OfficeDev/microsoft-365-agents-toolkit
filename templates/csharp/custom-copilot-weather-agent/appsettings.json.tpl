{
  "AgentApplicationOptions": {
    "StartTypingTimer": true,
    "RemoveRecipientMention": false,
    "NormalizeMentions": false
  },

  "TokenValidation": {
    "Audiences": [
      "{{ClientId}}" // this is the Client ID used for the Azure Bot
    ]
  },

  "Connections": {
    "BotServiceConnection": {
      "Assembly": "Microsoft.Agents.Authentication.Msal",
      "Type": "MsalAuth",
      "Settings": {
        "AuthType": "ClientSecret", // this is the AuthType for the connection, valid values can be found in Microsoft.Agents.Authentication.Msal.Model.AuthTypes.  The default is ClientSecret.
        "AuthorityEndpoint": "https://login.microsoftonline.com/botframework.com",
        "ClientId": "{{ClientId}}", // this is the Client ID used for the connection.
        "ClientSecret": "00000000-0000-0000-0000-000000000000", // this is the Client Secret used for the connection.
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
      "Microsoft.Copilot": "Warning",
      "Microsoft.Hosting.Lifetime": "Information"
    }
  },

  // This is the configuration for the AI services, use environeent variables or user secrets to store sensitive information.
  // Do not store sensitive information in this file
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
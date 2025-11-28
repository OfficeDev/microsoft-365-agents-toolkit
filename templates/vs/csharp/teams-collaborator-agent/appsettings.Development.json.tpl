{
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft.AspNetCore": "Warning"
    },
    "Microsoft.Teams": {
      "Enable": "*",
      "Level": "debug"
    }
  },
  "AllowedHosts": "*",
  "Teams": {
    "ClientId": "",
    "ClientSecret": "",
    "BotType": ""
  },
{{#useOpenAI}}
  "OpenAI": {
    "ApiKey": "",
    "DefaultModel": "gpt-3.5-turbo"
  },
{{/useOpenAI}}
{{#useAzureOpenAI}}
  "Azure": {
    "OpenAIApiKey": "",
    "OpenAIEndpoint": "",
    "OpenAIDeploymentName": "" 
  },
{{/useAzureOpenAI}}
  "Database": {
    "Type": "sqlite",
    "SqlitePath": "conversations.db"
  }
}

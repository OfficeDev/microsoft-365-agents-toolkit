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
    "ApiKey": "{{{originalOpenAIKey}}}",
    "Model": "gpt-3.5-turbo"
  }
{{/useOpenAI}}
{{#useAzureOpenAI}}
  "Azure": {
    "OpenAIApiKey": "{{{originalAzureOpenAIKey}}}",
    "OpenAIEndpoint": "{{{azureOpenAIEndpoint}}}",
    "OpenAIDeploymentName": "{{{azureOpenAIDeploymentName}}}" 
  },
{{/useAzureOpenAI}}
  "Database": {
    "Type": "sqlite",
    "SqlitePath": "conversations.db"
  }
}

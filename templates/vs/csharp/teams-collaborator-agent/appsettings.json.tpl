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
    "DefaultModel": "gpt-4o"
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
    "Type": "",
    "Password": ""
  }
}

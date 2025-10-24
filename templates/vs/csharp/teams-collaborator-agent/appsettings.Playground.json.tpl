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
    "ClientId": "${{BOT_ID}}",
    "ClientSecret": "${{SECRET_BOT_PASSWORD}}",
    "BotType": ""
  },
{{#useOpenAI}}
  "OpenAI": {
    "ApiKey": "${{SECRET_OPENAI_API_KEY}}",
    "DefaultModel": "${{OPENAI_MODEL_NAME}}"
  },
{{/useOpenAI}}
{{#useAzureOpenAI}}
  "Azure": {
    "OpenAIApiKey": "${{SECRET_AZURE_OPENAI_API_KEY}}",
    "OpenAIEndpoint": "${{AZURE_OPENAI_ENDPOINT}}",
    "OpenAIDeploymentName": "${{AZURE_OPENAI_DEPLOYMENT_NAME}}" 
  },
{{/useAzureOpenAI}}
  "Storage": {
    "Type": "sqlite",
    "ConnectionString": "Data Source=conversations.db"
  }
}

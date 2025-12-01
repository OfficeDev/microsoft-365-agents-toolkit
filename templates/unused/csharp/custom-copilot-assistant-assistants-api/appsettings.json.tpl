{
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft": "Information",
      "Microsoft.Hosting.Lifetime": "Information"
    }
  },
  "AllowedHosts": "*",
  "BOT_ID": "",
  "BOT_PASSWORD": "",
  "BOT_TYPE": "",
  "BOT_TENANT_ID": "",
{{#useOpenAI}}
  "OpenAI": {
    "ApiKey": "",
    "AssistantId": ""
  }
{{/useOpenAI}}
{{#useAzureOpenAI}}
  "Azure": {
    "OpenAIApiKey": "",
    "OpenAIEndpoint": "",
    "OpenAIDeploymentName": "",
    "OpenAIAssistantId": ""
  }
{{/useAzureOpenAI}}
}
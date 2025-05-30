{
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft": "Information",
      "Microsoft.Hosting.Lifetime": "Information",
      "Microsoft.Teams.AI": "Trace"
    }
  },
  "AllowedHosts": "*",
  "BOT_ID": "",
  "BOT_PASSWORD": "",
  "BOT_TYPE": "",
{{#useOpenAI}}
  "OpenAI": {
    "ApiKey": "",
    "EmbeddingModel": ""
  },
  "Azure": {
    "AISearchApiKey": "",
    "AISearchEndpoint": ""
  }
{{/useOpenAI}}
{{#useAzureOpenAI}}
  "Azure": {
    "OpenAIApiKey": "",
    "OpenAIEndpoint": "",
    "OpenAIDeploymentName": "",
    "OpenAIEmbeddingDeploymentName": "",
    "AISearchApiKey": "",
    "AISearchEndpoint": ""
  }
{{/useAzureOpenAI}}
}
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
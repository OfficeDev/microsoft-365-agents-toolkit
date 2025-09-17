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
    "EmbeddingModel": "{{{openAIEmbeddingModel}}}"
  },
  "Azure": {
    "AISearchApiKey": "{{{originalAzureAISearchApiKey}}}",
    "AISearchEndpoint": "{{{azureAISearchEndpoint}}}"
  }
{{/useOpenAI}}
{{#useAzureOpenAI}}
  "Azure": {
    "OpenAIApiKey": "{{{originalAzureOpenAIKey}}}",
    "OpenAIEndpoint": "{{{azureOpenAIEndpoint}}}",
    "OpenAIDeploymentName": "{{{azureOpenAIDeploymentName}}}",
    "OpenAIEmbeddingDeploymentName": "{{{azureOpenAIEmbeddingDeploymentName}}}",
    "AISearchApiKey": "{{{originalAzureAISearchApiKey}}}",
    "AISearchEndpoint": "{{{azureAISearchEndpoint}}}"
  }
{{/useAzureOpenAI}}
}
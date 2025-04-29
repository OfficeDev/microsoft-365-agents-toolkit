{
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft": "Warning",
      "Microsoft.Hosting.Lifetime": "Information"
    }
  },
  "AllowedHosts": "*",
  "BOT_ID": "",
  "BOT_PASSWORD": "",
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
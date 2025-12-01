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
    "AssistantId": ""
  }
{{/useOpenAI}}
{{#useAzureOpenAI}}
  "Azure": {
    "OpenAIApiKey": "{{{originalAzureOpenAIKey}}}",
    "OpenAIEndpoint": "{{{azureOpenAIEndpoint}}}",
    "OpenAIDeploymentName": "{{{azureOpenAIDeploymentName}}}",
    "OpenAIAssistantId": ""
  }
{{/useAzureOpenAI}}
}
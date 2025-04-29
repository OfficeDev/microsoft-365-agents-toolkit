{
  "Connections": {
    "BotServiceConnection": {
      "Settings": {
        "AuthType": "ClientSecret",
        "AuthorityEndpoint": "https://login.microsoftonline.com/botframework.com",
        "ClientId": "00000000-0000-0000-0000-000000000000",
        "ClientSecret": "00000000-0000-0000-0000-000000000000",
        "Scopes": [
          "https://api.botframework.com/.default"
        ]
      }
    }
  },
{{#useOpenAI}}
  "OpenAI": {
    "ApiKey": "{{{originalOpenAIKey}}}"
  }
{{/useOpenAI}}
{{#useAzureOpenAI}}
  "Azure": {
    "OpenAIApiKey": "{{{originalAzureOpenAIKey}}}",
    "OpenAIEndpoint": "{{{azureOpenAIEndpoint}}}",
    "OpenAIDeploymentName": "{{{azureOpenAIDeploymentName}}}" 
  }
{{/useAzureOpenAI}}
}
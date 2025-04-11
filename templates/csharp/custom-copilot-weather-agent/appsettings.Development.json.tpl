{
  "TokenValidation": {
    "Audiences": {
      "ClientId": ""
    }
  },
  "Connections": {
    "BotServiceConnection": {
      "Settings": {
        "ClientId": "",
        "ClientSecret": ""
      }
    }
  },
{{#useOpenAI}}
  "OpenAI": {
    "ApiKey": ""
  }
{{/useOpenAI}}
{{#useAzureOpenAI}}
  "Azure": {
    "OpenAIApiKey": "",
    "OpenAIEndpoint": "",
    "OpenAIDeploymentName": "" 
  }
{{/useAzureOpenAI}}
}
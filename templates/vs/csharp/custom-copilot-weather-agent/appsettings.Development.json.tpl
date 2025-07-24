{
  "TokenValidation": {
    "Audiences": {
      "ClientId": ""
    }
  },
  "Connections": {
    "ServiceConnection": {
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
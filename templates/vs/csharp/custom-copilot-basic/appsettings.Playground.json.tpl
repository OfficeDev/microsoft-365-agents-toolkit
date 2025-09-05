{
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft": "Warning",
      "Microsoft.Hosting.Lifetime": "Information"
    }
  },
  "AllowedHosts": "*",
   "Teams": {
		"ClientId": "",
		"ClientSecret": "",
    "TenantId": "",
		"BotType": ""
	}
{{#useOpenAI}}
  "OpenAI": {
    "ApiKey": ""，
    "Model": "gpt-3.5-turbo"
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
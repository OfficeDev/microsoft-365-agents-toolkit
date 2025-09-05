{
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft": "Information",
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
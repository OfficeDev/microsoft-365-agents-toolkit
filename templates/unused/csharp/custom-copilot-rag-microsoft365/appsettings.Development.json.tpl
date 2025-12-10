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
  "BOT_DOMAIN": "",
	"AAD_APP_CLIENT_ID": "",
	"AAD_APP_CLIENT_SECRET": "",
	"AAD_APP_TENANT_ID": "",
	"AAD_APP_OAUTH_AUTHORITY_HOST": "",
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
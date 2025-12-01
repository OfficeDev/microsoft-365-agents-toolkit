const config = {
  MicrosoftAppId: process.env.BOT_ID,
  MicrosoftAppType: process.env.BOT_TYPE,
  MicrosoftAppTenantId: process.env.BOT_TENANT_ID,
  MicrosoftAppPassword: process.env.BOT_PASSWORD,
  {{#useOpenAI}}
  openAIKey: process.env.OPENAI_API_KEY,
  openAIAssistantId: process.env.OPENAI_ASSISTANT_ID,
  {{/useOpenAI}}
  {{#useAzureOpenAI}}
  azureOpenAIKey: process.env.AZURE_OPENAI_API_KEY,
  azureOpenAIEndpoint: process.env.AZURE_OPENAI_ENDPOINT,
  azureOpenAIAssistantId: process.env.AZURE_OPENAI_ASSISTANT_ID,
  {{/useAzureOpenAI}}
};

export default config;

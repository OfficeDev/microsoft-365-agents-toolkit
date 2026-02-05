const config = {
  azureOpenAIKey: process.env.AZURE_OPENAI_API_KEY,
  azureOpenAIEndpoint: process.env.AZURE_OPENAI_ENDPOINT,
  azureOpenAIDeploymentName: process.env.AZURE_OPENAI_DEPLOYMENT_NAME,

  // Microsoft Foundry (Azure AI Projects) configuration
  foundryProjectEndpoint: process.env.FOUNDRY_PROJECT_ENDPOINT,
  foundryAgentName: process.env.FOUNDRY_AGENT_NAME,
};

export default config;

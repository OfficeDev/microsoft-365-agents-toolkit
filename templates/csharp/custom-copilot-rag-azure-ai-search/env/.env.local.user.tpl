# This file includes environment variables that will not be committed to git by default. You can set these environment variables in your CI/CD system for your project.

# Secrets. Keys prefixed with `SECRET_` will be masked in Microsoft 365 Agents Toolkit logs.
SECRET_BOT_PASSWORD=
{{#useOpenAI}}
SECRET_OPENAI_API_KEY={{{openAIKey}}}
OPENAI_EMBEDDING_MODEL={{{openAIEmbeddingModel}}}
SECRET_AI_SEARCH_API_KEY={{{azureAISearchApiKey}}}
AI_SEARCH_ENDPOINT={{{azureAISearchEndpoint}}}
{{/useOpenAI}}
{{#useAzureOpenAI}}
SECRET_AZURE_OPENAI_API_KEY={{{azureOpenAIKey}}}
AZURE_OPENAI_ENDPOINT={{{azureOpenAIEndpoint}}}
AZURE_OPENAI_DEPLOYMENT_NAME={{{azureOpenAIDeploymentName}}}
AZURE_OPENAI_EMBEDDING_DEPLOYMENT_NAME={{{azureOpenAIEmbeddingDeploymentName}}}
SECRET_AI_SEARCH_API_KEY={{{azureAISearchApiKey}}}
AI_SEARCH_ENDPOINT={{{azureAISearchEndpoint}}}
{{/useAzureOpenAI}}

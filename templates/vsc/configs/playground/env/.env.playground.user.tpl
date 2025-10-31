{{#useOpenAI}}
SECRET_OPENAI_API_KEY={{{openAIKey}}}
{{/useOpenAI}}

{{#useAzureOpenAI}}
SECRET_AZURE_OPENAI_API_KEY={{{azureOpenAIKey}}}
AZURE_OPENAI_ENDPOINT='{{{azureOpenAIEndpoint}}}'
AZURE_OPENAI_DEPLOYMENT_NAME='{{{azureOpenAIDeploymentName}}}'
{{/useAzureOpenAI}}
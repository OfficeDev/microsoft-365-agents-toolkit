# This file includes environment variables that will not be committed to git by default. You can set these environment variables in your CI/CD system for your project.

# If you're adding a secret value, add SECRET_ prefix to the name so Microsoft 365 Agents Toolkit can handle them properly
# Secrets. Keys prefixed with `SECRET_` will be masked in Microsoft 365 Agents Toolkit logs.
{{#azureOpenAIKey}}
SECRET_AZURE_OPENAI_API_KEY={{{azureOpenAIKey}}}
{{/azureOpenAIKey}}
{{^azureOpenAIKey}}
SECRET_AZURE_OPENAI_API_KEY=
{{/azureOpenAIKey}}
{{#azureOpenAIEndpoint}}
AZURE_OPENAI_ENDPOINT='{{{azureOpenAIEndpoint}}}'
{{/azureOpenAIEndpoint}}
{{^azureOpenAIEndpoint}}
AZURE_OPENAI_ENDPOINT=
{{/azureOpenAIEndpoint}}
{{#azureOpenAIDeploymentName}}
AZURE_OPENAI_DEPLOYMENT_NAME='{{{azureOpenAIDeploymentName}}}'
{{/azureOpenAIDeploymentName}}
{{^azureOpenAIDeploymentName}}
AZURE_OPENAI_DEPLOYMENT_NAME=
{{/azureOpenAIDeploymentName}}
# Optional: Azure OpenAI API version; defaults in code if unset
{{#azureOpenAIApiVersion}}
AZURE_OPENAI_API_VERSION='{{{azureOpenAIApiVersion}}}'
{{/azureOpenAIApiVersion}}
{{^azureOpenAIApiVersion}}
AZURE_OPENAI_API_VERSION=
{{/azureOpenAIApiVersion}}
# Optional: OpenAI model ID when using OpenAI (non-Azure)
{{#openAIModel}}
OPENAI_MODEL='{{{openAIModel}}}'
{{/openAIModel}}
{{^openAIModel}}
OPENAI_MODEL=
{{/openAIModel}}
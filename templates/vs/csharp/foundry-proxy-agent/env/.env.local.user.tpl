# This file includes environment variables that will not be committed to git by default. You can set these environment variables in your CI/CD system for your project.

# Secrets. Keys prefixed with `SECRET_` will be masked in Microsoft 365 Agents Toolkit logs.
SECRET_BOT_PASSWORD=

# Azure AI Foundry Configuration (Required - fill in your values)
AZURE_AI_FOUNDRY_PROJECT_ENDPOINT={{{FoundryEndpoint}}}
AGENT_ID={{{FoundryAgentId}}}

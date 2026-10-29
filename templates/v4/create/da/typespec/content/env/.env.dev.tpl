# This file includes environment variables that will be committed to git by default.

# Built-in environment variables
TEAMSFX_ENV=dev
APP_NAME_SUFFIX=dev
GITHUB_API_URL=https://api.github.com
AGENT_SCOPE=shared

# Generated during provision, you can also add your own variables.
{{#NewDeveloperPortalApis}}
APP_ID=
{{/NewDeveloperPortalApis}}
{{^NewDeveloperPortalApis}}
TEAMS_APP_ID=
{{/NewDeveloperPortalApis}}

# Environment variables for Agent Evaluations
AZURE_AI_OPENAI_ENDPOINT=
AZURE_AI_API_KEY=
AZURE_AI_API_VERSION=
AZURE_AI_MODEL_NAME=
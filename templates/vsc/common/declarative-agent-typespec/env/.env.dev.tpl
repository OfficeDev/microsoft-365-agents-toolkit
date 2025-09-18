# This file includes environment variables that will be committed to git by default.

# Built-in environment variables
TEAMSFX_ENV=dev
APP_NAME_SUFFIX=dev
{{#ShareEnabled}}
AGENT_SCOPE=shared
{{/ShareEnabled}}

# Generated during provision, you can also add your own variables.
TEAMS_APP_ID=
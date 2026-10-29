# This file includes environment variables that will be committed to git by default.

# Built-in environment variables
TEAMSFX_ENV=local
APP_NAME_SUFFIX=local
AGENT_SCOPE=shared

# Generated during provision, you can also add your own variables.
{{#NewDeveloperPortalApis}}
APP_ID=
{{/NewDeveloperPortalApis}}
{{^NewDeveloperPortalApis}}
TEAMS_APP_ID=
{{/NewDeveloperPortalApis}}

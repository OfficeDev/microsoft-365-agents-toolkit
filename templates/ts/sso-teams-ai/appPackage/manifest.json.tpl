{
  "$schema": "https://developer.microsoft.com/json-schemas/teams/v1.20/MicrosoftTeams.schema.json",
  "version": "1.0.0",
  "manifestVersion": "1.20",
  "id": "${{TEAMS_APP_ID}}",
  "name": {
    "short": "{{appName}}-${{APP_NAME_SUFFIX}}",
    "full": "{{appName}}"
  },
  "developer": {
    "name": "Microsoft",
    "mpnId": "",
    "websiteUrl": "https://microsoft.com",
    "privacyUrl": "https://privacy.microsoft.com/privacystatement",
    "termsOfUseUrl": "https://www.microsoft.com/legal/terms-of-use"
  },
  "description": {
    "short": "Sample bot that uses MS Graph",
    "full": "Sample bot that uses MS Graph"
  },
  "icons": {
    "outline": "outline.png",
    "color": "color.png"
  },
  "accentColor": "#FFFFFF",
  "staticTabs": [
    {
      "entityId": "conversations",
      "scopes": ["personal"]
    },
    {
      "entityId": "about",
      "scopes": ["personal"]
    }
  ],
  "bots": [
    {
      "botId": "${{BOT_ID}}",
      "scopes": ["personal", "team", "groupChat"],
      "isNotificationOnly": false,
      "supportsCalling": false,
      "supportsVideo": false,
      "supportsFiles": false
    }
  ],
  "validDomains": [
    "${{BOT_DOMAIN}}",
    "*.botframework.com"
  ],
  "permissions": ["identity"],
  "webApplicationInfo": {
    "id": "${{BOT_ID}}",
    "resource": "api://botid-${{BOT_ID}}"
  }
}

{
{{#CEAEnabled}} 
    "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/vdevPreview/MicrosoftTeams.schema.json",
    "manifestVersion": "devPreview",
    "version": "1.0.0",
{{/CEAEnabled}}
{{^CEAEnabled}} 
    "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.19/MicrosoftTeams.schema.json",
    "manifestVersion": "1.19",
    "version": "1.0.0",
{{/CEAEnabled}}
    "id": "${{TEAMS_APP_ID}}",
    "developer": {
        "name": "Teams App, Inc.",
        "websiteUrl": "https://www.example.com",
        "privacyUrl": "https://www.example.com/privacy",
        "termsOfUseUrl": "https://www.example.com/termofuse"
    },
    "icons": {
        "color": "color.png",
        "outline": "outline.png"
    },
    "name": {
        "short": "{{appName}}${{APP_NAME_SUFFIX}}",
        "full": "full name for {{appName}}"
    },
    "description": {
        "short": "Short description of {{appName}}",
        "full": "Full description of {{appName}}"
    },
    "accentColor": "#FFFFFF",
        {{#CEAEnabled}} 
    "copilotAgents": {
        "customEngineAgents": [
            {
                "type": "bot",
                "id": "${{BOT_ID}}"
            }
        ]
    },
    {{/CEAEnabled}}
    "bots": [
        {
            "botId": "${{BOT_ID}}",
            "scopes": [
                "copilot",
                "personal",
                "team",
                "groupChat"
            ],
            "supportsFiles": false,
            "isNotificationOnly": false,
            "commandLists": [
                {
                    "scopes": [
                        "copilot",
                        "personal"
                    ],
                    "commands": [
                        {
                            "title": "Solve the equation: 3x + 11= 14",
                            "description": "Help me solve the equation: 3x + 11= 14"
                        },
                        {
                            "title": "The weather of San Francisco",
                            "description": "The weather of San Francisco"
                        }
                    ]
                }
            ]
        }
    ],
    "composeExtensions": [
    ],
    "configurableTabs": [],
    "staticTabs": [],
    "permissions": [
        "identity",
        "messageTeamMembers"
    ],
    "validDomains": []
}
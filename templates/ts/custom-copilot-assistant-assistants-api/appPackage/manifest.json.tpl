{
    {{#LaunchAgentForTeamsInCopilotEnabled}} 
    "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/vdevPreview/MicrosoftTeams.schema.json",
    "manifestVersion": "devPreview",
    "version": "1.0.0",
    {{/LaunchAgentForTeamsInCopilotEnabled}}
    {{^LaunchAgentForTeamsInCopilotEnabled}} 
    "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.19/MicrosoftTeams.schema.json",
    "manifestVersion": "1.19",
    "version": "1.0.0",
    {{/LaunchAgentForTeamsInCopilotEnabled}}
    "id": "${{TEAMS_APP_ID}}",
    "developer": {
        "name": "My App, Inc.",
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
        "short": "short description for {{appName}}",
        "full": "full description for {{appName}}"
    },
    "accentColor": "#FFFFFF",
    {{#LaunchAgentForTeamsInCopilotEnabled}} 
    "copilotAgents": {
        "customEngineAgents": [
            {
                "type": "bot",
                "id": "${{BOT_ID}}"
            }
        ]
    },
    {{/LaunchAgentForTeamsInCopilotEnabled}}
    "bots": [
        {
            "botId": "${{BOT_ID}}",
            "scopes": [
                {{#LaunchAgentForTeamsInCopilotEnabled}} 
                "copilot",
                {{/LaunchAgentForTeamsInCopilotEnabled}}
                "personal"
            ],
            "supportsFiles": false,
            "isNotificationOnly": false,
            "commandLists": [
                {
                    "scopes": [
                        {{#LaunchAgentForTeamsInCopilotEnabled}} 
                        "copilot",
                        {{/LaunchAgentForTeamsInCopilotEnabled}}
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
    "composeExtensions": [],
    "configurableTabs": [],
    "staticTabs": [],
    "permissions": [
        "identity"
    ],
    "validDomains": []
}
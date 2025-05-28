{
    "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.21/MicrosoftTeams.schema.json",
    "manifestVersion": "1.21",
    "version": "1.0.0",
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
                {{#CEAEnabled}} 
                "copilot",
                {{/CEAEnabled}}
                "personal",
                "team",
                "groupChat"
            ],
            "supportsFiles": false,
            "isNotificationOnly": false,
            "commandLists": [
                {
                    "scopes": [
                        {{#CEAEnabled}} 
                        "copilot",
                        {{/CEAEnabled}}
                        "personal"
                    ],
                    "commands": [
                        {
                            "title": "List Contoso history in table",
                            "description": "Tell me the history of Contoso Electronics, format in a table."
                        },
                        {
                            "title": "Compare Contoso Electronics plan",
                            "description": "Compare different Contoso Electronics benefit package plans"
                        },
                        {
                            "title": "Summarize PerksPlus Program",
                            "description": "Summarize Contoso Electronics PerksPlus Program"
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
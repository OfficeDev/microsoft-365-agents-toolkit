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
        "short": "short description for {{appName}}",
        "full": "full description for {{appName}}"
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
                            "title": "Create task:remind drink tonight",
                            "description": "Create a task for me to remind me drink water tonight"
                        },
                        {
                            "title": "Delete all my current tasks",
                            "description": "Delete all my current tasks"
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
        "identity",
        "messageTeamMembers"
    ],
    "validDomains": []
}

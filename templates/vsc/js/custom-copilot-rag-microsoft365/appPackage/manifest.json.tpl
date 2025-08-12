{
    "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.22/MicrosoftTeams.schema.json",
    "manifestVersion": "1.22",
    "version": "1.0.0",
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
    "bots": [
        {
            "botId": "${{BOT_ID}}",
            "scopes": [
                "team",
                "groupChat",
                "personal"
            ],
            "supportsFiles": false,
            "isNotificationOnly": false,
            "commandLists": [
                {
                    "scopes": [
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
    "composeExtensions": [],
    "configurableTabs": [],
    "staticTabs": [],
    "permissions": [
        "identity",
        "messageTeamMembers"
    ],
    "validDomains": [
        "${{BOT_DOMAIN}}"
    ],
    "webApplicationInfo": {
        "id": "${{AAD_APP_CLIENT_ID}}",
        "resource": "api://botid-${{BOT_ID}}"
    }
}
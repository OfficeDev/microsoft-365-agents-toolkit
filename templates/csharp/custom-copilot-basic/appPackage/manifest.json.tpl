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
                "personal",
                "team",
                "groupChat"
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
                            "title": "How can you help me?",
                            "description": "How can you help me?"
                        },
                        {
                            "title": "How to develop agent for Teams?",
                            "description": "How can I develop apps with Microsoft 365 Agents Toolkit?"
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
{
    "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.23/MicrosoftTeams.schema.json",
    "manifestVersion": "1.23",
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
        "short": "Teams Collaborator Agent - Summarize chats, track action items, and search conversations",
        "full": "An intelligent agent that helps you collaborate more effectively in Teams by summarizing conversations, identifying action items, and searching through chat history."
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
                {{^CEAEnabled}}
                "team",
                "groupChat",
                {{/CEAEnabled}}
                "personal"
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
                            "title": "Summarize this chat",
                            "description": "Generate a summary of the current conversation"
                        },
                        {
                            "title": "Show action items",
                            "description": "List all action items from this conversation"
                        },
                        {
                            "title": "Search conversations",
                            "description": "Search through conversation history"
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
    "authorization": {
        "permissions": {
            "resourceSpecific": [
                {
                    "name": "ChatMessage.Read.Chat",
                    "type": "Application"
                }
            ]
        }
    },
    "validDomains": []
}
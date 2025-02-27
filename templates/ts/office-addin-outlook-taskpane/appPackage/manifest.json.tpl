{
    "$schema": "https://raw.githubusercontent.com/OfficeDev/microsoft-teams-app-schema/preview/DevPreview/MicrosoftTeams.schema.json",
    "manifestVersion": "devPreview",
    "version": "1.0.1",
    "id": "{{manifestId}}",
    "packageName": "com.microsoft.teams.extension",
    "developer": {
        "name": "Contoso App, Inc.",
        "websiteUrl": "https://test",
        "privacyUrl": "https://test/index.html#/privacy",
        "termsOfUseUrl": "https://test/index.html#/termsofuse"
    },
    "icons": {
        "color": "color.png",
        "outline": "outline.png"
    },
    "name": {
        "short": "{{appName}}",
        "full": "Full name for {{appName}}"
    },
    "description": {
        "short": "Short description of Contoso",
        "full": "Full description of Contoso"
    },
    "accentColor": "#FFFFFF",
    "bots": [],
    "composeExtensions": [],
    "configurableTabs": [
        {
            "configurationUrl": "https://test/index.html#/config",
            "canUpdateConfiguration": true,
            "scopes": [
                "team",
                "groupChat"
            ]
        }
    ],
    "staticTabs": [
        {
            "entityId": "index0",
            "name": "Personal Tab",
            "contentUrl": "https://test/index.html#/tab",
            "websiteUrl": "https://test/index.html#/tab",
            "scopes": [
                "personal"
            ]
        }
    ],
    "permissions": [
        "identity",
        "messageTeamMembers"
    ],
    "authorization": {
        "permissions": {
            "resourceSpecific": [
                {
                    "name": "MailboxItem.Read.User",
                    "type": "Delegated"
                }
            ]
        }
    },
    "validDomains": [
        "test.com",
        "contoso.com"
    ],
    "extensions": [
        {
            "requirements": {
                "scopes": [
                    "mail"
                ],
                "capabilities": [
                    {
                        "name": "Mailbox",
                        "minVersion": "1.3"
                    }
                ]
            },
            "runtimes": [
                {
                    "requirements": {
                        "capabilities": [
                            {
                                "name": "Mailbox",
                                "minVersion": "1.3"
                            }
                        ]
                    },
                    "id": "TaskPaneRuntime",
                    "type": "general",
                    "code": {
                        "page": "https://localhost:53000/taskpane.html"
                    },
                    "lifetime": "short",
                    "actions": [
                        {
                            "id": "TaskPaneRuntimeShow",
                            "type": "openPage",
                            "pinnable": false,
                            "view": "dashboard"
                        }
                    ]
                },
                {
                    "id": "CommandsRuntime",
                    "type": "general",
                    "code": {
                        "page": "https://localhost:53000/commands.html",
                        "script": "https://localhost:53000/commands.js"
                    },
                    "lifetime": "short",
                    "actions": [
                        {
                            "id": "action",
                            "type": "executeFunction",
                            "displayName": "action"
                        }
                    ]
                }
            ],
            "ribbons": [
                {
                    "contexts": [
                        "mailRead"
                    ],
                    "tabs": [
                        {
                            "builtInTabId": "TabDefault",
                            "groups": [
                                {
                                    "id": "msgReadGroup",
                                    "label": "Contoso Add-in",
                                    "icons": [
                                        {
                                            "size": 16,
                                            "url": "https://localhost:53000/assets/icon-16.png"
                                        },
                                        {
                                            "size": 32,
                                            "url": "https://localhost:53000/assets/icon-32.png"
                                        },
                                        {
                                            "size": 80,
                                            "url": "https://localhost:53000/assets/icon-80.png"
                                        }
                                    ],
                                    "controls": [
                                        {
                                            "id": "msgReadOpenPaneButton",
                                            "type": "button",
                                            "label": "Show Taskpane",
                                            "icons": [
                                                {
                                                    "size": 16,
                                                    "url": "https://localhost:53000/assets/icon-16.png"
                                                },
                                                {
                                                    "size": 32,
                                                    "url": "https://localhost:53000/assets/icon-32.png"
                                                },
                                                {
                                                    "size": 80,
                                                    "url": "https://localhost:53000/assets/icon-80.png"
                                                }
                                            ],
                                            "supertip": {
                                                "title": "Show Taskpane",
                                                "description": "Opens a pane displaying all available properties."
                                            },
                                            "actionId": "TaskPaneRuntimeShow"
                                        },
                                        {
                                            "id": "ActionButton",
                                            "type": "button",
                                            "label": "Perform an action",
                                            "icons": [
                                                {
                                                    "size": 16,
                                                    "url": "https://localhost:53000/assets/icon-16.png"
                                                },
                                                {
                                                    "size": 32,
                                                    "url": "https://localhost:53000/assets/icon-32.png"
                                                },
                                                {
                                                    "size": 80,
                                                    "url": "https://localhost:53000/assets/icon-80.png"
                                                }
                                            ],
                                            "supertip": {
                                                "title": "Perform an action",
                                                "description": "Perform an action when clicked."
                                            },
                                            "actionId": "action"
                                        }
                                    ]
                                }
                            ]
                        }
                    ]
                }
            ]
        }
    ]
}
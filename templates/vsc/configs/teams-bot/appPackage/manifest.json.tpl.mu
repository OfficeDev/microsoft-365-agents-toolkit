{
    <<#CEA>>
    "copilotAgents": {
        "customEngineAgents": [
            {
                "type": "bot",
                "id": "${{BOT_ID}}"
            }
        ]
    },
    <</CEA>>
    "bots": [
        {
            "botId": "${{BOT_ID}}",
            "scopes": [
                <<#CEA>>
                "copilot",
                <</CEA>>
                "personal",
                "team",
                "groupChat"
            ],
            "supportsFiles": false,
            "isNotificationOnly": false,
            "commandLists": [
                {
                  "scopes": [
                    <<#CEA>>
                    "copilot",
                    <</CEA>>
                    "personal",
                    "team",
                    "groupChat"
                ],
                  "commands": [
                      {
                          "title": "Hi",
                          "description": "Say hi to the bot."
                      }
                  ]
                }
            ]
        }
    ],
    "permissions": [
        "identity",
        "messageTeamMembers"
    ]
}
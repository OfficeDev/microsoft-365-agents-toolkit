{
    "staticTabs": [
        {
            "entityId": "index0",
            "name": "Home",
            "contentUrl": "${{APP_ENDPOINT}}/tabs/home",
            "websiteUrl": "${{APP_ENDPOINT}}/tabs/home",
            "scopes": [
                "personal",
                "groupChat",
                "team"
            ]
        }
    ],
    "permissions": [
        "identity",
        "messageTeamMembers"
    ],
    "validDomains": [
        "${{APP_DOMAIN}}"
    ]
}
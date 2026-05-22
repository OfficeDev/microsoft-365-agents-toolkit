{
    "version": "0.2.0",
    "configurations": [
        {
            "name": "Preview App (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://admin.teams.microsoft.com/policies/manage-apps",
            "presentation": {
                "group": "remote",
                "order": 1
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Preview App (Chrome)",
            "type": "chrome",
            "request": "launch",
            "url": "https://admin.teams.microsoft.com/policies/manage-apps",
            "presentation": {
                "group": "remote",
                "order": 2
            },
            "internalConsoleOptions": "neverOpen"
        }
    ]
}

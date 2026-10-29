{
    "version": "0.2.0",
    "configurations": [
        {
            "name": "Preview in Teams (Edge)",
            "type": "msedge",
            "request": "launch",
            {{#NewDeveloperPortalApis}}
            "url": "https://teams.microsoft.com/l/app/${{APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
            {{/NewDeveloperPortalApis}}
            {{^NewDeveloperPortalApis}}
            "url": "https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
            {{/NewDeveloperPortalApis}}
            "presentation": {
                "group": "remote",
                "order": 1
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Preview in Teams (Chrome)",
            "type": "chrome",
            "request": "launch",
            {{#NewDeveloperPortalApis}}
            "url": "https://teams.microsoft.com/l/app/${{APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
            {{/NewDeveloperPortalApis}}
            {{^NewDeveloperPortalApis}}
            "url": "https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
            {{/NewDeveloperPortalApis}}
            "presentation": {
                "group": "remote",
                "order": 2
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Preview in Teams (Desktop)",
            "type": "node",
            "request": "launch",
            "preLaunchTask": "Start App in Desktop Client (Remote)",
            "presentation": {
                "group": "remote",
                "order": 3
            },
            "internalConsoleOptions": "neverOpen"
        }
    ]
}

{
    "version": "0.2.0",
    "configurations": [
        {
            "name": "Launch Remote (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
            "presentation": {
                "group": "3-remote",
                "order": 1
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Launch Remote (Chrome)",
            "type": "chrome",
            "request": "launch",
            "url": "https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
            "presentation": {
                "group": "3-remote",
                "order": 2
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Launch App (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://teams.microsoft.com/l/app/${{local:TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
            "cascadeTerminateToConfigurations": [
                "Python: Remote Attach"
            ],
            "presentation": {
                "group": "all",
                "hidden": true
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Launch App (Chrome)",
            "type": "chrome",
            "request": "launch",
            "url": "https://teams.microsoft.com/l/app/${{local:TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
            "cascadeTerminateToConfigurations": [
                "Python: Remote Attach"
            ],
            "presentation": {
                "group": "all",
                "hidden": true
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Python: Remote Attach",
            "type": "python",
            "request": "attach",
            "connect": {
                "host": "localhost",
                "port": 5678
            },
            "pathMappings": [
                {
                    "localRoot": "${workspaceFolder}",
                    "remoteRoot": "."
                }
            ],
            "presentation": {
                "group": "all",
                "hidden": true
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Launch Remote (Desktop)",
            "type": "python",
            "request": "launch",
            "preLaunchTask": "Start App in Desktop Client (Remote)",
            "presentation": {
                "group": "3-remote",
                "order": 3
            },
            "internalConsoleOptions": "neverOpen",
        }
    ],
    "compounds": [
        {
            "name": "Debug in Teams (Edge)",
            "configurations": [
                "Launch App (Edge)",
                "Python: Remote Attach"
            ],
            "preLaunchTask": "Start App Locally",
            "presentation": {
{{#enableTestToolByDefault}}
                "group": "2-local",
{{/enableTestToolByDefault}}
{{^enableTestToolByDefault}}
                "group": "1-local",
{{/enableTestToolByDefault}}
                "order": 1
            },
            "stopAll": true
        },
        {
            "name": "Debug in Teams (Chrome)",
            "configurations": [
                "Launch App (Chrome)",
                "Python: Remote Attach"
            ],
            "preLaunchTask": "Start App Locally",
            "presentation": {
{{#enableTestToolByDefault}}
                "group": "2-local",
{{/enableTestToolByDefault}}
{{^enableTestToolByDefault}}
                "group": "1-local",
{{/enableTestToolByDefault}}
                "order": 2
            },
            "stopAll": true
        },
        {
            "name": "Debug in Teams (Desktop)",
            "configurations": [
                "Python: Remote Attach"
            ],
            "preLaunchTask": "Start App in Desktop Client",
            "presentation": {
                "group": "2-local",
                "order": 3
            },
            "stopAll": true
        },
        {
            "name": "Debug in Microsoft 365 Agents Playground",
            "configurations": [
                "Python: Remote Attach"
            ],
            "preLaunchTask": "Start App in Microsoft 365 Agents Playground",
            "presentation": {
{{#enableTestToolByDefault}}
                "group": "1-local",
{{/enableTestToolByDefault}}
{{^enableTestToolByDefault}}
                "group": "2-local",
{{/enableTestToolByDefault}}
                "order": 1
            },
            "stopAll": true
        }
    ]
}

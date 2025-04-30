{
    "version": "0.2.0",
    "configurations": [
        {
            "name": "Launch Remote in Teams (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
            "presentation": {
                "group": "group 1: Teams",
                "order": 4
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Launch Remote in Teams (Chrome)",
            "type": "chrome",
            "request": "launch",
            "url": "https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
            "presentation": {
                "group": "group 1: Teams",
                "order": 5
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Launch Remote in Outlook (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://outlook.office.com/mail?${account-hint}",
            "presentation": {
                "group": "group 2: Outlook",
                "order": 3
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Launch Remote in Outlook (Chrome)",
            "type": "chrome",
            "request": "launch",
            "url": "https://outlook.office.com/mail?${account-hint}",
            "presentation": {
                "group": "group 2: Outlook",
                "order": 3
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Launch App in Teams (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://teams.microsoft.com/l/app/${{local:TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
            "cascadeTerminateToConfigurations": [
                "Attach to Local Service"
            ],
            "presentation": {
                "group": "all",
                "hidden": true
            },
            "internalConsoleOptions": "neverOpen",
            "perScriptSourcemaps": "yes"
        },
        {
            "name": "Launch App in Teams (Chrome)",
            "type": "chrome",
            "request": "launch",
            "url": "https://teams.microsoft.com/l/app/${{local:TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
            "cascadeTerminateToConfigurations": [
                "Attach to Local Service"
            ],
            "presentation": {
                "group": "all",
                "hidden": true
            },
            "internalConsoleOptions": "neverOpen",
            "perScriptSourcemaps": "yes"
        },
        {
            "name": "Launch App in Outlook (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://outlook.office.com/mail?${account-hint}",
            "cascadeTerminateToConfigurations": [
                "Attach to Local Service"
            ],
            "presentation": {
                "group": "all",
                "hidden": true
            },
            "internalConsoleOptions": "neverOpen",
            "perScriptSourcemaps": "yes"
        },
        {
            "name": "Launch App in Outlook (Chrome)",
            "type": "chrome",
            "request": "launch",
            "url": "https://outlook.office.com/mail?${account-hint}",
            "cascadeTerminateToConfigurations": [
                "Attach to Local Service"
            ],
            "presentation": {
                "group": "all",
                "hidden": true
            },
            "internalConsoleOptions": "neverOpen",
            "perScriptSourcemaps": "yes"
        },
        {
            "name": "Attach to Local Service",
            "type": "node",
            "request": "attach",
            "port": 9239,
            "restart": true,
            "presentation": {
                "group": "all",
                "hidden": true
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Launch Remote in Teams (Desktop)",
            "type": "node",
            "request": "launch",
            "preLaunchTask": "Start App in Desktop Client (Remote)",
            "presentation": {
                "group": "group 1: Teams",
                "order": 6
            },
            "internalConsoleOptions": "neverOpen",
        }
    ],
    "compounds": [
        {
            "name": "Debug in Microsoft 365 Agents Playground",
            "configurations": [
                "Attach to Local Service"
            ],
            "preLaunchTask": "Start App in Microsoft 365 Agents Playground",
            "presentation": {
{{#enableMETestToolByDefault}}
                "group": "group 0: Microsoft 365 Agents Playground",
{{/enableMETestToolByDefault}}
{{^enableMETestToolByDefault}}
                "group": "group 3: Microsoft 365 Agents Playground",
{{/enableMETestToolByDefault}}
                "order": 1
            },
            "stopAll": true
        },
        {
            "name": "Debug in Teams (Edge)",
            "configurations": [
                "Launch App in Teams (Edge)",
                "Attach to Local Service"
            ],
            "preLaunchTask": "Start App Locally",
            "presentation": {
                "group": "group 1: Teams",
                "order": 1
            },
            "stopAll": true
        },
        {
            "name": "Debug in Teams (Chrome)",
            "configurations": [
                "Launch App in Teams (Chrome)",
                "Attach to Local Service"
            ],
            "preLaunchTask": "Start App Locally",
            "presentation": {
                "group": "group 1: Teams",
                "order": 2
            },
            "stopAll": true
        },
        {
            "name": "Debug in Outlook (Edge)",
            "configurations": [
                "Launch App in Outlook (Edge)",
                "Attach to Local Service"
            ],
            "preLaunchTask": "Start App Locally",
            "presentation": {
                "group": "group 2: Outlook",
                "order": 1
            },
            "stopAll": true
        },
        {
            "name": "Debug in Teams (Desktop)",
            "configurations": [
                "Attach to Local Service"
            ],
            "preLaunchTask": "Start App in Desktop Client",
            "presentation": {
                "group": "group 1: Teams",
                "order": 3
            },
            "stopAll": true
        },
        {
            "name": "Debug in Outlook (Chrome)",
            "configurations": [
                "Launch App in Outlook (Chrome)",
                "Attach to Local Service"
            ],
            "preLaunchTask": "Start App Locally",
            "presentation": {
                "group": "group 2: Outlook",
                "order": 2
            },
            "stopAll": true
        }
    ]
}
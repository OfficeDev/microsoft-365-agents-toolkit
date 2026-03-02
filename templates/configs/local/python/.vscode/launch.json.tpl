{
    "version": "0.2.0",
    "configurations": [
        {
            "name": "Launch App in Teams (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://teams.microsoft.com/l/app/$\{{local:TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
            "presentation": {
                "group": "all",
                "hidden": true
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Launch App in Teams (Chrome)",
            "type": "chrome",
            "request": "launch",
            "url": "https://teams.microsoft.com/l/app/$\{{local:TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
            "presentation": {
                "group": "all",
                "hidden": true
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Start Python",
            "type": "debugpy",
            "request": "launch",
            "program": "${workspaceFolder}/src/app.py",
            "cwd": "${workspaceFolder}/src",
            "console": "integratedTerminal"
        },
{{#supportCopilot}}
{{/supportCopilot}}
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
        }
    ],
    "compounds": [
        {
            "name": "Debug in Teams (Edge)",
            "configurations": [
                "Launch App in Teams (Edge)",
                "Start Python"
            ],
            "cascadeTerminateToConfigurations": ["Start Python"],
            "preLaunchTask": "Start App Locally",
            "presentation": {
                "group": "2-local",
                "order": 1
            },
            "stopAll": true
        },
        {
            "name": "Debug in Teams (Chrome)",
            "configurations": [
                "Launch App in Teams (Chrome)",
                "Start Python"
            ],
            "cascadeTerminateToConfigurations": ["Start Python"],
            "preLaunchTask": "Start App Locally",
            "presentation": {
                "group": "2-local",
                "order": 2
            },
            "stopAll": true
        },
        {
            "name": "Debug in Teams (Desktop)",
            "configurations": [
                "Start Python"
            ],
            "preLaunchTask": "Start App in Desktop Client",
            "presentation": {
                "group": "2-local",
                "order": 3
            },
            "stopAll": true
        },
{{#supportCopilot}}
{{/supportCopilot}}
    ]
}
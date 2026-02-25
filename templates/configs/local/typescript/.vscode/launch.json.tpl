{
    "version": "0.2.0",
    "configurations": [
        {
            "name": "Launch App in Teams (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://teams.microsoft.com/l/app/$\{{local:TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
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
            "url": "https://teams.microsoft.com/l/app/$\{{local:TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
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
{{#supportCopilot}}
        {
            "name": "Launch in Copilot (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://m365.cloud.microsoft/chat/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429/$\{local:agent-hint}?auth=2&${account-hint}&developerMode=Basic",
            "cascadeTerminateToConfigurations": ["Attach to Local Service"],
            "presentation": {
                "group": "all",
                "hidden": true
            },
            "internalConsoleOptions": "neverOpen",
            "runtimeArgs": [
                "--remote-debugging-port=9222",
                "--no-first-run"
            ]
        },
        {
            "name": "Launch in Copilot (Chrome)",
            "type": "chrome",
            "request": "launch",
            "url": "https://m365.cloud.microsoft/chat/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429/$\{local:agent-hint}?auth=2&${account-hint}&developerMode=Basic",
            "cascadeTerminateToConfigurations": ["Attach to Local Service"],
            "presentation": {
                "group": "all",
                "hidden": true
            },
            "internalConsoleOptions": "neverOpen",
            "runtimeArgs": [
                "--remote-debugging-port=9223",
                "--no-first-run"
            ]
        },
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
                "Attach to Local Service"
            ],
            "preLaunchTask": "Start App Locally",
            "presentation": {
                "group": "2-local",
                "order": 4
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
                "group": "2-local",
                "order": 5
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
                "group": "2-local",
                "order": 6
            },
            "stopAll": true
        },
{{#supportCopilot}}
        {
            "name": "(Preview) Debug in Copilot (Edge)",
            "configurations": [
                "Launch in Copilot (Edge)",
                "Attach to Local Service"
            ],
            "preLaunchTask": "Start App Locally",
            "presentation": {
                "group": "2-local",
                "order": 1
            },
            "stopAll": true
        },
        {
            "name": "(Preview) Debug in Copilot (Chrome)",
            "configurations": [
                "Launch in Copilot (Chrome)",
                "Attach to Local Service"
            ],
            "preLaunchTask": "Start App Locally",
            "presentation": {
                "group": "2-local",
                "order": 2
            },
            "stopAll": true
        },
{{/supportCopilot}}
    ]
}
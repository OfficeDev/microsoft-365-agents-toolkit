{{^DeclarativeCopilot}}
{
    "version": "0.2.0",
    "configurations": [
        {
            "name": "Launch App in the Microsoft 365 app (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://www.office.com/chat?auth=2&developerMode=Basic",
            "cascadeTerminateToConfigurations": [
                "Attach to Backend"
            ],
            "presentation": {
                "group": "all",
                "hidden": true
            },
            "internalConsoleOptions": "neverOpen",
            "perScriptSourcemaps": "yes",
            "runtimeArgs": [
                "--remote-debugging-port=9222",
                "--no-first-run",
                "--user-data-dir=${env:TEMP}/copilot-edge-user-data-dir"
            ]
        },
        {
            "name": "Launch App in the Microsoft 365 app (Chrome)",
            "type": "chrome",
            "request": "launch",
            "url": "https://www.office.com/chat?auth=2&developerMode=Basic",
            "cascadeTerminateToConfigurations": [
                "Attach to Backend"
            ],
            "presentation": {
                "group": "all",
                "hidden": true
            },
            "internalConsoleOptions": "neverOpen",
            "perScriptSourcemaps": "yes",
            "runtimeArgs": [
                "--remote-debugging-port=9223",
                "--no-first-run",
                "--user-data-dir=${env:TEMP}/copilot-chrome-user-data-dir"
            ]
        },
        {
            "name": "Launch App in Teams (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://teams.microsoft.com?${account-hint}",
            "cascadeTerminateToConfigurations": [
                "Attach to Backend"
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
            "url": "https://teams.microsoft.com?${account-hint}",
            "cascadeTerminateToConfigurations": [
                "Attach to Backend"
            ],
            "presentation": {
                "group": "all",
                "hidden": true
            },
            "internalConsoleOptions": "neverOpen",
            "perScriptSourcemaps": "yes"
        },
        {
            "name": "Preview in Teams (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://teams.microsoft.com?${account-hint}",
            "presentation": {
                "group": "group 2: Teams",
                "order": 4
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Preview in Teams (Chrome)",
            "type": "chrome",
            "request": "launch",
            "url": "https://teams.microsoft.com?${account-hint}",
            "presentation": {
                "group": "group 2: Teams",
                "order": 5
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Preview in Teams (Desktop)",
            "type": "node",
            "request": "launch",
            "preLaunchTask": "Start desktop client",
            "presentation": {
                "group": "group 2: Teams",
                "order": 6
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Attach to Backend",
            "type": "node",
            "request": "attach",
            "port": 9229,
            "restart": true,
            "presentation": {
                "group": "all",
                "hidden": true
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Preview in the Microsoft 365 app (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://www.office.com/chat?auth=2&developerMode=Basic",
            "presentation": {
                "group": "group 1: the Microsoft 365 app",
                "order": 3
            },
            "internalConsoleOptions": "neverOpen",
            "runtimeArgs": [
                "--remote-debugging-port=9224",
                "--no-first-run",
                "--user-data-dir=${env:TEMP}/copilot-edge-user-data-dir"
            ]
        },
        {
            "name": "Preview in the Microsoft 365 app (Chrome)",
            "type": "chrome",
            "request": "launch",
            "url": "https://www.office.com/chat?auth=2&developerMode=Basic",
            "presentation": {
                "group": "group 1: the Microsoft 365 app",
                "order": 4
            },
            "internalConsoleOptions": "neverOpen",
            "runtimeArgs": [
                "--remote-debugging-port=9225",
                "--no-first-run",
                "--user-data-dir=${env:TEMP}/copilot-chrome-user-data-dir"
            ]
        }
    ],
    "compounds": [
        {
            "name": "Debug in the Microsoft 365 app (Edge)",
            "configurations": [
                "Launch App in the Microsoft 365 app (Edge)",
                "Attach to Backend"
            ],
            "preLaunchTask": "Start Teams App Locally",
            "presentation": {
                "group": "group 1: the Microsoft 365 app",
                "order": 1
            },
            "stopAll": true
        },
        {
            "name": "Debug in the Microsoft 365 app (Chrome)",
            "configurations": [
                "Launch App in the Microsoft 365 app (Chrome)",
                "Attach to Backend"
            ],
            "preLaunchTask": "Start Teams App Locally",
            "presentation": {
                "group": "group 1: the Microsoft 365 app",
                "order": 2
            },
            "stopAll": true
        },
        {
            "name": "Debug in Teams (Edge)",
            "configurations": [
                "Launch App in Teams (Edge)",
                "Attach to Backend"
            ],
            "preLaunchTask": "Start Teams App Locally",
            "presentation": {
                "group": "group 2: Teams",
                "order": 1
            },
            "stopAll": true
        },
        {
            "name": "Debug in Teams (Chrome)",
            "configurations": [
                "Launch App in Teams (Chrome)",
                "Attach to Backend"
            ],
            "preLaunchTask": "Start Teams App Locally",
            "presentation": {
                "group": "group 2: Teams",
                "order": 2
            },
            "stopAll": true
        },
        {
            "name": "Debug in Teams (Desktop)",
            "configurations": [
                "Attach to Backend"
            ],
            "preLaunchTask": "Start Teams App in Desktop Client",
            "presentation": {
                "group": "group 2: Teams",
                "order": 3
            },
            "stopAll": true
        }
    ]
}
{{/DeclarativeCopilot}}
{{#DeclarativeCopilot}}
{
    "version": "0.2.0",
    "configurations": [
        {
            "name": "Launch App in Teams (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://m365.cloud.microsoft/chat/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429/${local:agent-hint}?auth=2&developerMode=Basic",
            "cascadeTerminateToConfigurations": [
                "Attach to Backend"
            ],
            "presentation": {
                "group": "all",
                "hidden": true
            },
            "internalConsoleOptions": "neverOpen",
            "perScriptSourcemaps": "yes",
            "runtimeArgs": [
                "--remote-debugging-port=9222",
                "--no-first-run",
                "--user-data-dir=${env:TEMP}/copilot-msedge-user-data-dir"
            ]
        },
        {
            "name": "Launch App in Teams (Chrome)",
            "type": "chrome",
            "request": "launch",
            "url": "https://m365.cloud.microsoft/chat/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429/${local:agent-hint}?auth=2&developerMode=Basic",
            "cascadeTerminateToConfigurations": [
                "Attach to Backend"
            ],
            "presentation": {
                "group": "all",
                "hidden": true
            },
            "internalConsoleOptions": "neverOpen",
            "perScriptSourcemaps": "yes",
            "runtimeArgs": [
                "--remote-debugging-port=9223",
                "--no-first-run",
                "--user-data-dir=${env:TEMP}/copilot-chrome-user-data-dir"
            ]
        },
        {
            "name": "Preview in Copilot (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://m365.cloud.microsoft/chat/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429/${agent-hint}?auth=2&developerMode=Basic",
            "presentation": {
                "group": "remote",
                "order": 1
            },
            "internalConsoleOptions": "neverOpen",
            "runtimeArgs": [
                "--remote-debugging-port=9224",
                "--no-first-run",
                "--user-data-dir=${env:TEMP}/copilot-msedge-user-data-dir"
            ]
        },
        {
            "name": "Preview in Copilot (Chrome)",
            "type": "chrome",
            "request": "launch",
            "url": "https://m365.cloud.microsoft/chat/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429/${agent-hint}?auth=2&developerMode=Basic",
            "presentation": {
                "group": "remote",
                "order": 2
            },
            "internalConsoleOptions": "neverOpen",
            "runtimeArgs": [
                "--remote-debugging-port=9225",
                "--no-first-run",
                "--user-data-dir=${env:TEMP}/copilot-chrome-user-data-dir"
            ]
        },
        {
            "name": "Attach to Backend",
            "type": "node",
            "request": "attach",
            "port": 9229,
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
            "name": "Debug in Copilot (Edge)",
            "configurations": [
                "Launch App in Teams (Edge)",
                "Attach to Backend"
            ],
            "preLaunchTask": "Start Teams App Locally",
            "presentation": {
                "group": "all",
                "order": 1
            },
            "stopAll": true
        },
        {
            "name": "Debug in Copilot (Chrome)",
            "configurations": [
                "Launch App in Teams (Chrome)",
                "Attach to Backend"
            ],
            "preLaunchTask": "Start Teams App Locally",
            "presentation": {
                "group": "all",
                "order": 2
            },
            "stopAll": true
        }
    ]
}
{{/DeclarativeCopilot}}
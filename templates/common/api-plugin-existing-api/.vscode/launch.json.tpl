{{^DeclarativeCopilot}}
{
    "version": "0.2.0",
    "configurations": [
        {
            "name": "Preview in the Microsoft 365 app (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://www.office.com/chat?auth=2&developerMode=Basic",
            "presentation": {
                "group": "group 1: the Microsoft 365 app",
                "order": 1
            },
            "internalConsoleOptions": "neverOpen",
            "runtimeArgs": [
                "--remote-debugging-port=9222",
                "--no-first-run",
                "--user-data-dir=${env:TEMP}/copilot-msedge-user-data-dir"
            ]
        },
        {
            "name": "Preview in the Microsoft 365 app (Chrome)",
            "type": "chrome",
            "request": "launch",
            "url": "https://www.office.com/chat?auth=2&developerMode=Basic",
            "presentation": {
                "group": "group 1: the Microsoft 365 app",
                "order": 2
            },
            "internalConsoleOptions": "neverOpen",
            "runtimeArgs": [
                "--remote-debugging-port=9223",
                "--no-first-run",
                "--user-data-dir=${env:TEMP}/copilot-chrome-user-data-dir"
            ]
        },
        {
            "name": "Preview in Teams (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://teams.microsoft.com?${account-hint}",
            "presentation": {
                "group": "group 2: Teams",
                "order": 1
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
                "order": 2
            },
            "internalConsoleOptions": "neverOpen"
        }
    ]
}
{{/DeclarativeCopilot}}
{{#DeclarativeCopilot}}
{
    "version": "0.2.0",
    "configurations": [
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
                "--remote-debugging-port=9222",
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
                "--remote-debugging-port=9223",
                "--no-first-run",
                "--user-data-dir=${env:TEMP}/copilot-chrome-user-data-dir"
            ]
        }
    ]
}
{{/DeclarativeCopilot}}


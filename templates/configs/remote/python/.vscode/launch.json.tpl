{
    "version": "0.2.0",
    "configurations": [
{{#supportCopilot}}
        {
            "name": "(Preview) View Remote Agent in Copilot (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://m365.cloud.microsoft/chat/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429/${agent-hint}?auth=2&${account-hint}&developerMode=Basic",
            "presentation": {
                "group": "3-remote",
                "order": 1
            },
            "internalConsoleOptions": "neverOpen",
            "runtimeArgs": [
                "--remote-debugging-port=9222",
                "--no-first-run"
            ]
        },
        {
            "name": "(Preview) View Remote Agent in Copilot (Chrome)",
            "type": "chrome",
            "request": "launch",
            "url": "https://m365.cloud.microsoft/chat/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429/${agent-hint}?auth=2&${account-hint}&developerMode=Basic",
            "presentation": {
                "group": "3-remote",
                "order": 2
            },
            "internalConsoleOptions": "neverOpen",
            "runtimeArgs": [
                "--remote-debugging-port=9223",
                "--no-first-run"
            ]
        },
{{/supportCopilot}}
        {
            "name": "View Remote App in Teams (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://teams.microsoft.com/l/app/$\{{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
            "presentation": {
                "group": "3-remote",
                "order": 4
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "View Remote App in Teams (Chrome)",
            "type": "chrome",
            "request": "launch",
            "url": "https://teams.microsoft.com/l/app/$\{{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
            "presentation": {
                "group": "3-remote",
                "order": 5
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "View Remote App in Teams (Desktop)",
            "type": "node",
            "request": "launch",
            "preLaunchTask": "Start App in Desktop Client (Remote)",
            "presentation": {
                "group": "3-remote",
                "order": 6
            },
            "internalConsoleOptions": "neverOpen",
        }
    ]
}
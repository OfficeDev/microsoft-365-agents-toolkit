{
  "version": "0.2.0",
  "configurations": [
    {
      "name": "Launch Remote (Edge)",
      "type": "msedge",
      "request": "launch",
      "url": "https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
      "presentation": {
        "group": "1-Teams",
        "order": 4
      },
      "internalConsoleOptions": "neverOpen"
    },
    {
      "name": "Launch Remote (Chrome)",
      "type": "chrome",
      "request": "launch",
      "url": "https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
      "presentation": {
        "group": "1-Teams",
        "order": 5
      },
      "internalConsoleOptions": "neverOpen"
    },
    {
      "name": "Launch Remote (Desktop)",
      "type": "node",
      "request": "launch",
      "preLaunchTask": "Start Teams App in Desktop Client (Remote)",
      "presentation": {
        "group": "1-Teams",
        "order": 6
      },
      "internalConsoleOptions": "neverOpen"
    },
    {
      "name": "Launch App (Edge)",
      "type": "msedge",
      "request": "launch",
      "url": "https://teams.microsoft.com/l/app/${{local:TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
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
    {
        "name": "Start Microsoft 365 Agents Playground",
        "type": "node",
        "request": "launch",
        "program": "${workspaceFolder}/devTools/teamsapptester/node_modules/@microsoft/teams-app-test-tool/cli.js",
        "args": [
            "start"
        ],
        "env": {
          "PATH": "${workspaceFolder}/devTools/nodejs{{pathDelimiter}}${env:PATH}"
        },
        "cwd": "${workspaceFolder}",
        "console": "integratedTerminal",
        "internalConsoleOptions": "neverOpen"
    {{#CEAEnabled}}
    },
    {
      "name": "Launch Remote in Copilot (Edge)",
      "type": "msedge",
      "request": "launch",
      "url": "https://m365.cloud.microsoft/chat/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429/${agent-hint}?auth=2&${account-hint}&developerMode=Basic",
      "presentation": {
        "group": "3-M365",
        "order": 3
      },
      "internalConsoleOptions": "neverOpen",
      "runtimeArgs": [
          "--remote-debugging-port=9222",
          "--no-first-run"
      ]
    },
    {
      "name": "Launch Remote in Copilot (Chrome)",
      "type": "chrome",
      "request": "launch",
      "url": "https://m365.cloud.microsoft/chat/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429/${agent-hint}?auth=2&${account-hint}&developerMode=Basic",
      "presentation": {
        "group": "3-M365",
        "order": 4
      },
      "internalConsoleOptions": "neverOpen",
      "runtimeArgs": [
          "--remote-debugging-port=9223",
          "--no-first-run"
      ]
    },
    {
      "name": "Local debug in Copilot (Edge)",
      "type": "msedge",
      "request": "launch",
      "url": "https://m365.cloud.microsoft/chat/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429/${local:agent-hint}?auth=2&${account-hint}&developerMode=Basic",
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
      "name": "Local debug in Copilot (Chrome)",
      "type": "chrome",
      "request": "launch",
      "url": "https://m365.cloud.microsoft/chat/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429/${local:agent-hint}?auth=2&${account-hint}&developerMode=Basic",
      "presentation": {
        "group": "all",
        "hidden": true
      },
      "internalConsoleOptions": "neverOpen",
      "runtimeArgs": [
          "--remote-debugging-port=9223",
          "--no-first-run"
      ]
    {{/CEAEnabled}}
    }
  ],
  "compounds": [
    {
      "name": "Debug in Teams (Edge)",
      "configurations": ["Launch App (Edge)", "Start Python"],
      "cascadeTerminateToConfigurations": ["Start Python"],
      "preLaunchTask": "Start Teams App Locally",
      "presentation": {
        "group": "1-Teams",
        "order": 1
      },
      "stopAll": true
    },
    {
      "name": "Debug in Teams (Chrome)",
      "configurations": ["Launch App (Chrome)", "Start Python"],
      "cascadeTerminateToConfigurations": ["Start Python"],
      "preLaunchTask": "Start Teams App Locally",
      "presentation": {
        "group": "1-Teams",
        "order": 2
      },
      "stopAll": true
    },
    {
      "name": "Debug in Teams (Desktop)",
      "configurations": ["Start Python"],
      "preLaunchTask": "Start Teams App in Desktop Client",
      "presentation": {
        "group": "1-Teams",
        "order": 3
      },
      "stopAll": true
    },
    {
        "name": "Debug in Microsoft 365 Agents Playground",
        "configurations": [
            "Start Python",
            "Start Microsoft 365 Agents Playground"
        ],
        "cascadeTerminateToConfigurations": [
            "Start Microsoft 365 Agents Playground"
        ],
        "preLaunchTask": "Deploy (Microsoft 365 Agents Playground)",
        "presentation": {
            "group": "0-TestTool",
            "order": 1
        },
        "stopAll": true
    {{#CEAEnabled}}
    },
    {
      "name": "Debug in Copilot (Edge)",
      "configurations": ["Local debug in Copilot (Edge)", "Start Python"],
      "cascadeTerminateToConfigurations": ["Start Python"],
      "preLaunchTask": "Start Teams App Locally",
      "presentation": {
        "group": "3-M365",
        "order": 1
      },
      "stopAll": true
    },
    {
      "name": "Debug in Copilot (Chrome)",
      "configurations": ["Local debug in Copilot (Chrome)", "Start Python"],
      "cascadeTerminateToConfigurations": ["Start Python"],
      "preLaunchTask": "Start Teams App Locally",
      "presentation": {
        "group": "3-M365",
        "order": 2
      },
      "stopAll": true
    {{/CEAEnabled}}
    }
  ]
}

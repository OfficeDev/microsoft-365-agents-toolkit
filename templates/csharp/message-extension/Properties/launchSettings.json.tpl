{
  "profiles": {
{{^isNewProjectTypeEnabled}}
{{#enableTestToolByDefault}}
    // Debug project within Microsoft 365 Agents Playground
    "Microsoft 365 Agents Playground (browser)": {
      "commandName": "Project",
      "dotnetRunMessages": true,
      "launchBrowser": true,
      "launchTestTool": true,
      "launchUrl": "http://localhost:56150",
      "applicationUrl": "http://localhost:5130",
      "environmentVariables": {
        "ASPNETCORE_ENVIRONMENT": "TestTool",
        "TEAMSFX_NOTIFICATION_STORE_FILENAME": ".notification.testtoolstore.json",
        "UPDATE_TEAMS_APP": "false"
      },
      "hotReloadProfile": "aspnetcore"
    },
{{/enableTestToolByDefault}}
    // Debug project within Teams
    "Microsoft Teams (browser)": {
      "commandName": "Project",
      "dotnetRunMessages": true,
      "launchBrowser": true,
      "launchUrl": "https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&appTenantId=${{TEAMS_APP_TENANT_ID}}&login_hint=${{TEAMSFX_M365_USER_NAME}}",
      "applicationUrl": "https://localhost:7130;http://localhost:5130",
      "environmentVariables": {
        "ASPNETCORE_ENVIRONMENT": "Development"
      },
      "hotReloadProfile": "aspnetcore"
    },
    "Outlook (browser)": {
      "commandName": "Project",
      "dotnetRunMessages": true,
      "launchBrowser": true,
      "launchUrl": "https://outlook.office.com/mail?login_hint=${{TEAMSFX_M365_USER_NAME}}",
      "applicationUrl": "http://localhost:5130",
      "environmentVariables": {
        "ASPNETCORE_ENVIRONMENT": "Development"
      },
      "hotReloadProfile": "aspnetcore"
    },
{{^enableTestToolByDefault}}
    // Debug project within Microsoft 365 Agents Playground
    "Microsoft 365 Agents Playground (browser)": {
      "commandName": "Project",
      "dotnetRunMessages": true,
      "launchBrowser": true,
      "launchTestTool": true,
      "launchUrl": "http://localhost:56150",
      "applicationUrl": "http://localhost:5130",
      "environmentVariables": {
        "ASPNETCORE_ENVIRONMENT": "TestTool",
        "TEAMSFX_NOTIFICATION_STORE_FILENAME": ".notification.testtoolstore.json",
        "UPDATE_TEAMS_APP": "false"
      },
      "hotReloadProfile": "aspnetcore"
    },
{{/enableTestToolByDefault}}
    //// Uncomment following profile to debug project only (without launching Teams)
    //,
    //"Start Project (not in Teams)": {
    //  "commandName": "Project",
    //  "dotnetRunMessages": true,
    //  "launchBrowser": true,
    //  "applicationUrl": "https://localhost:7130;http://localhost:5130",
    //  "environmentVariables": {
    //    "ASPNETCORE_ENVIRONMENT": "Development"
    //  },
    //  "hotReloadProfile": "aspnetcore"
    //}
{{/isNewProjectTypeEnabled}}
{{#isNewProjectTypeEnabled}}
{{#enableTestToolByDefault}}
    "Microsoft 365 Agents Playground": {
      "commandName": "Project",
      "dotnetRunMessages": true,
      "applicationUrl": "http://localhost:5130",
      "environmentVariables": {
        "ASPNETCORE_ENVIRONMENT": "TestTool",
        "TEAMSFX_NOTIFICATION_STORE_FILENAME": ".notification.testtoolstore.json",
        "UPDATE_TEAMS_APP": "false"
      },
      "hotReloadProfile": "aspnetcore"
    },
{{/enableTestToolByDefault}}
    "Start Project": {
      "commandName": "Project",
      "dotnetRunMessages": true,
      "applicationUrl": "https://localhost:7130;http://localhost:5130",
      "environmentVariables": {
        "ASPNETCORE_ENVIRONMENT": "Development"
      },
      "hotReloadProfile": "aspnetcore"
    },
{{^enableTestToolByDefault}}
    "Microsoft 365 Agents Playground": {
      "commandName": "Project",
      "dotnetRunMessages": true,
      "applicationUrl": "http://localhost:5130",
      "environmentVariables": {
        "ASPNETCORE_ENVIRONMENT": "TestTool",
        "TEAMSFX_NOTIFICATION_STORE_FILENAME": ".notification.testtoolstore.json",
        "UPDATE_TEAMS_APP": "false"
      },
      "hotReloadProfile": "aspnetcore"
    },
{{/enableTestToolByDefault}}
{{/isNewProjectTypeEnabled}}
  }
}
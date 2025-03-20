{
  "profiles": {
{{^isNewProjectTypeEnabled}}
{{#enableTestToolByDefault}}
    // Debug project within Teams App Test Tool
    "Teams App Test Tool (browser)": {
      "commandName": "Project",
      "commandLineArgs": "host start --port 5130 --pause-on-error",
      "dotnetRunMessages": true,
      "launchBrowser": true,
      "launchTestTool": true,
      "launchUrl": "http://localhost:56150",
      "environmentVariables": {
        "AZURE_FUNCTIONS_ENVIRONMENT": "TestTool",
        "TEAMSFX_NOTIFICATION_STORE_FILENAME": ".notification.testtoolstore.json",
        "TEAMSFX_NOTIFICATION_LOCALSTORE_DIR": "../../..", // Path to project folder $(MSBuildProjectDirectory)
        "UPDATE_TEAMS_APP": "false"
      }
    },
{{/enableTestToolByDefault}}
    // Debug project within Teams
    "Microsoft Teams (browser)": {
      "commandName": "Project",
      "commandLineArgs": "host start --port 5130 --pause-on-error",
      "dotnetRunMessages": true,
      "launchBrowser": true,
      "launchUrl": "https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&appTenantId=${{TEAMS_APP_TENANT_ID}}&login_hint=${{TEAMSFX_M365_USER_NAME}}",
      "environmentVariables": {
        "ASPNETCORE_ENVIRONMENT": "Development",
        "TEAMSFX_NOTIFICATION_LOCALSTORE_DIR": "../../.." // Path to project folder $(MSBuildProjectDirectory)
      }
    },
{{^enableTestToolByDefault}}
    // Debug project within Teams App Test Tool
    "Teams App Test Tool (browser)": {
      "commandName": "Project",
      "commandLineArgs": "host start --port 5130 --pause-on-error",
      "dotnetRunMessages": true,
      "launchBrowser": true,
      "launchTestTool": true,
      "launchUrl": "http://localhost:56150",
      "environmentVariables": {
        "AZURE_FUNCTIONS_ENVIRONMENT": "TestTool",
        "TEAMSFX_NOTIFICATION_STORE_FILENAME": ".notification.testtoolstore.json",
        "TEAMSFX_NOTIFICATION_LOCALSTORE_DIR": "../../..", // Path to project folder $(MSBuildProjectDirectory)
        "UPDATE_TEAMS_APP": "false"
      }
    },
{{/enableTestToolByDefault}}
    //// Uncomment following profile to debug project only (without launching Teams)
    //,
    //"Start Project (not in Teams)": {
    //  "commandName": "Project",
    //  "commandLineArgs": "host start --port 5130 --pause-on-error",
    //  "dotnetRunMessages": true,
    //  "environmentVariables": {
    //    "ASPNETCORE_ENVIRONMENT": "Development",
    //    "TEAMSFX_NOTIFICATION_LOCALSTORE_DIR": "../../.." // Path to project folder $(MSBuildProjectDirectory)
    //  }
    //}
{{/isNewProjectTypeEnabled}}
{{#isNewProjectTypeEnabled}}
{{#enableTestToolByDefault}}
    // Launch project with TestTool environment, will be used by Teams App Test Tool
    "Teams App Test Tool": {
      "commandName": "Project",
      "commandLineArgs": "host start --port 5130 --pause-on-error",
      "dotnetRunMessages": true,
      "environmentVariables": {
        "AZURE_FUNCTIONS_ENVIRONMENT": "TestTool",
        "TEAMSFX_NOTIFICATION_STORE_FILENAME": ".notification.testtoolstore.json",
        // Path to project folder $(MSBuildProjectDirectory), used in Microsoft.TeamsFx package.
        "TEAMSFX_NOTIFICATION_LOCALSTORE_DIR": "../../..", 
        "UPDATE_TEAMS_APP": "false"
      }
    },
{{/enableTestToolByDefault}}
    // Launch project directly
    "Start Project": {
      "commandName": "Project",
      "commandLineArgs": "host start --port 5130 --pause-on-error",
      "dotnetRunMessages": true,
      "environmentVariables": {
        "ASPNETCORE_ENVIRONMENT": "Development",
        // Path to project folder $(MSBuildProjectDirectory), used in Microsoft.TeamsFx package.
        "TEAMSFX_NOTIFICATION_LOCALSTORE_DIR": "../../.."
      }
    },
{{^enableTestToolByDefault}}
    // Launch project with TestTool environment, will be used by Teams App Test Tool
    "Teams App Test Tool": {
      "commandName": "Project",
      "commandLineArgs": "host start --port 5130 --pause-on-error",
      "dotnetRunMessages": true,
      "environmentVariables": {
        "AZURE_FUNCTIONS_ENVIRONMENT": "TestTool",
        "TEAMSFX_NOTIFICATION_STORE_FILENAME": ".notification.testtoolstore.json",
        // Path to project folder $(MSBuildProjectDirectory), used in Microsoft.TeamsFx package.
        "TEAMSFX_NOTIFICATION_LOCALSTORE_DIR": "../../..", 
        "UPDATE_TEAMS_APP": "false"
      }
    },
{{/enableTestToolByDefault}}
{{/isNewProjectTypeEnabled}}
  }
}
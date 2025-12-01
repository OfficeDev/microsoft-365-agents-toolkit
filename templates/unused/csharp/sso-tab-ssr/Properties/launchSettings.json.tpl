{
  "profiles": {
{{^isNewProjectTypeEnabled}}
    // Debug project within Teams
    "Microsoft Teams (browser)": {
      "commandName": "Project",
      "dotnetRunMessages": true,
      "launchBrowser": true,
      "launchUrl": "https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&appTenantId=${{TEAMS_APP_TENANT_ID}}&login_hint=${{TEAMSFX_M365_USER_NAME}}",
      "applicationUrl": "https://localhost:44302;http://localhost:2544",
      "environmentVariables": {
        "ASPNETCORE_ENVIRONMENT": "Development"
      },
      "hotReloadProfile": "aspnetcore"
    },
    "Microsoft 365 app (browser)": {
      "commandName": "Project",
      "dotnetRunMessages": true,
      "launchBrowser": true,
      "launchUrl": "https://www.office.com/m365apps/${{M365_APP_ID}}?auth=2&login_hint=${{TEAMSFX_M365_USER_NAME}}",
      "applicationUrl": "https://localhost:44302;http://localhost:2544",
      "environmentVariables": {
        "ASPNETCORE_ENVIRONMENT": "Development"
      },
      "hotReloadProfile": "aspnetcore"
    },
    "Outlook (browser)": {
      "commandName": "Project",
      "dotnetRunMessages": true,
      "launchBrowser": true,
      "launchUrl": "https://outlook.office.com/host/${{M365_APP_ID}}?login_hint=${{TEAMSFX_M365_USER_NAME}}",
      "applicationUrl": "https://localhost:44302;http://localhost:2544",
      "environmentVariables": {
        "ASPNETCORE_ENVIRONMENT": "Development"
      },
      "hotReloadProfile": "aspnetcore"
    }
    //// Uncomment following profile to debug project only (without launching Teams)
    //,
    //"Start Project (not in Teams)": {
    //  "commandName": "Project",
    //  "dotnetRunMessages": true,
    //  "launchBrowser": true,
    //  "applicationUrl": "https://localhost:44302;http://localhost:2544",
    //  "environmentVariables": {
    //    "ASPNETCORE_ENVIRONMENT": "Development"
    //  },
    //  "hotReloadProfile": "aspnetcore"
    //}
{{/isNewProjectTypeEnabled}}
{{#isNewProjectTypeEnabled}}
    "Start Project": {
      "commandName": "Project",
      "dotnetRunMessages": true,
      "applicationUrl": "https://localhost:44302;http://localhost:2544",
      "environmentVariables": {
        "ASPNETCORE_ENVIRONMENT": "Development"
      },
      "hotReloadProfile": "aspnetcore"
    }
{{/isNewProjectTypeEnabled}}
  }
}
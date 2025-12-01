{
  "profiles": {
{{^isNewProjectTypeEnabled}}
    "Microsoft Teams (browser)": {
      "commandName": "Project",
      "commandLineArgs": "host start --port 5130 --pause-on-error",
      "dotnetRunMessages": true,
      "launchBrowser": true,
      "launchUrl": "https://teams.microsoft.com?appTenantId=${{TEAMS_APP_TENANT_ID}}&login_hint=${{TEAMSFX_M365_USER_NAME}}",
      "environmentVariables": {
        "ASPNETCORE_ENVIRONMENT": "Development"
      },
      "hotReloadProfile": "aspnetcore"
    }
    //// Uncomment following profile to debug project only (without launching Teams)
    //,
    //"Start Project (not in Teams)": {
    //  "commandName": "Project",
    //  "commandLineArgs": "host start --port 5130 --pause-on-error",
    //  "dotnetRunMessages": true,
    //  "applicationUrl": "https://localhost:7130;http://localhost:5130",
    //  "environmentVariables": {
    //    "ASPNETCORE_ENVIRONMENT": "Development"
    //  },
    //  "hotReloadProfile": "aspnetcore"
    //}
{{/isNewProjectTypeEnabled}}
{{#isNewProjectTypeEnabled}}
    "Start Project": {
      "commandName": "Project",
      "commandLineArgs": "host start --port 5130 --pause-on-error",
      "dotnetRunMessages": true,
      "environmentVariables": {
        "ASPNETCORE_ENVIRONMENT": "Development"
      },
      "hotReloadProfile": "aspnetcore"
    }
{{/isNewProjectTypeEnabled}}
  }
}

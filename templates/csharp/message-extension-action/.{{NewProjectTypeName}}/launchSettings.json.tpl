{
  "profiles": {
{{#enableTestToolByDefault}}
    // Launch project within Microsoft 365 Agents playground
    "Microsoft 365 Agents playground (browser)": {
      "commandName": "Project",
      "environmentVariables": { "UPDATE_TEAMS_APP": "false", "DEFAULT_CHANNEL_ID": "emulator" },
      "launchTestTool": true,
      "launchUrl": "http://localhost:56150",
    },
{{/enableTestToolByDefault}}
    // Debug project within Teams
    "Microsoft Teams (browser)": {
      "commandName": "Project",
      "launchUrl": "https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&appTenantId=${{TEAMS_APP_TENANT_ID}}&login_hint=${{TEAMSFX_M365_USER_NAME}}",
    },
    // Launch project within Teams without prepare Teams App dependencies
    "Microsoft Teams (browser) (skip update app)": {
      "commandName": "Project",
      "environmentVariables": { "UPDATE_TEAMS_APP": "false" },
      "launchUrl": "https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&appTenantId=${{TEAMS_APP_TENANT_ID}}&login_hint=${{TEAMSFX_M365_USER_NAME}}"
    },
{{^enableTestToolByDefault}}
    // Launch project within Microsoft 365 Agents playground
    "Microsoft 365 Agents playground (browser)": {
      "commandName": "Project",
      "environmentVariables": { "UPDATE_TEAMS_APP": "false", "DEFAULT_CHANNEL_ID": "emulator" },
      "launchTestTool": true,
      "launchUrl": "http://localhost:56150",
    },
{{/enableTestToolByDefault}}
  }
}
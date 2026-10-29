{
  "profiles": {
    // Debug project within Teams
    "Microsoft Teams (browser)": {
      "commandName": "Project",
      {{#NewDeveloperPortalApis}}
      "launchUrl": "https://teams.microsoft.com/l/app/${{APP_ID}}?installAppPackage=true&webjoin=true&appTenantId=${{TEAMS_APP_TENANT_ID}}&login_hint=${{TEAMSFX_M365_USER_NAME}}"
      {{/NewDeveloperPortalApis}}
      {{^NewDeveloperPortalApis}}
      "launchUrl": "https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&appTenantId=${{TEAMS_APP_TENANT_ID}}&login_hint=${{TEAMSFX_M365_USER_NAME}}"
      {{/NewDeveloperPortalApis}}
    }
  }
}
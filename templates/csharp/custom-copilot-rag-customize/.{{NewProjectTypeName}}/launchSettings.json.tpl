{
  "profiles": {
{{#enableTestToolByDefault}}
    // Launch project within Teams App Test Tool
    "Teams App Test Tool (browser)": {
      "commandName": "Project",
      "launchTestTool": true,
      "launchUrl": "http://localhost:56150",
    },
{{/enableTestToolByDefault}}
    // Launch project within Teams
    "Microsoft Teams (browser)": {
      "commandName": "Project",
      "launchUrl": "https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&appTenantId=${{TEAMS_APP_TENANT_ID}}&login_hint=${{TEAMSFX_M365_USER_NAME}}",
    },
{{^enableTestToolByDefault}}
    // Launch project within Teams App Test Tool
    "Teams App Test Tool (browser)": {
      "commandName": "Project",
      "launchTestTool": true,
      "launchUrl": "http://localhost:56150",
    },
{{/enableTestToolByDefault}}
{{#CEAEnabled}}
    // Launch project within M365 Copilot
    "Microsoft 365 Copilot (browser)": {
      "commandName": "Project",
      "launchUrl": "https://m365.cloud.microsoft/chat/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429/${{AGENT_HINT}}?auth=2"
    },
{{/CEAEnabled}}
  }
}
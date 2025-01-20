{
  "profiles": {
    {{^DeclarativeCopilot}}
    // Launch project within the Microsoft 365 app
    "Microsoft 365 app (browser)": {
      "commandName": "Project",
      "launchUrl": "https://www.office.com/chat?auth=2",
    },
    // Launch project within Teams
    "Microsoft Teams (browser)": {
      "commandName": "Project",
      "launchUrl": "https://teams.microsoft.com?appTenantId=${{TEAMS_APP_TENANT_ID}}&login_hint=${{TEAMSFX_M365_USER_NAME}}",
    }
    {{/DeclarativeCopilot}}
    {{#DeclarativeCopilot}}
    // Launch project within Copilot
    "Copilot (browser)": {
      "commandName": "Project",
      "launchUrl": "https://m365.cloud.microsoft/chat/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429/${{AGENT_HINT}}?auth=2"
    }
    {{/DeclarativeCopilot}}
  }
}
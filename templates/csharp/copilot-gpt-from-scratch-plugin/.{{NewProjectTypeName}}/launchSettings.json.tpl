{
  "profiles": {
    // Launch project within Copilot
    "Copilot (browser)": {
      "commandName": "Project",
      "launchUrl": "https://www.office.com/chat?auth=2",
    },
    // Launch project within Copilot without prepare Teams App dependencies
    "Copilot (browser) (skip update Teams App)": {
      "commandName": "Project",
      "environmentVariables": { "UPDATE_TEAMS_APP": "false" },
      "launchUrl": "https://www.office.com/chat?auth=2",
    }
  }
}
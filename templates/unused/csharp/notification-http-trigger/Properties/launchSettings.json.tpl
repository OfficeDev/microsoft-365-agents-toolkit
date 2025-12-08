{
  "profiles": {
    // Launch project directly
    "Start Project": {
      "commandName": "Project",
      "commandLineArgs": "host start --port 5130 --pause-on-error",
      "dotnetRunMessages": true,
      "environmentVariables": {
        "DOTNET_ENVIRONMENT": "Development",
        "ASPNETCORE_ENVIRONMENT": "Development",
        // Path to project folder $(MSBuildProjectDirectory), used in Microsoft.TeamsFx package.
        "TEAMSFX_NOTIFICATION_LOCALSTORE_DIR": "../../.."
      }
    },
  }
}
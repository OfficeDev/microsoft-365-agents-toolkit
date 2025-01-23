# yaml-language-server: $schema=https://aka.ms/teams-toolkit/v1.7/yaml.schema.json
# Visit https://aka.ms/teamsfx-v5.0-guide for details on this file
# Visit https://aka.ms/teamsfx-actions for details on actions
version: v1.7

provision:
  # Creates a Teams app
  - uses: teamsApp/create
    with:
      # Teams app name
      name: {{appName}}${{APP_NAME_SUFFIX}}
    # Write the information of created resources into environment file for
    # the specified environment variable(s).
    writeToEnvironmentFile:
      teamsAppId: TEAMS_APP_ID

  # Set OPENAPI_SERVER_URL for local launch
  - uses: script
    with:
      run:
        echo "::set-teamsfx-env OPENAPI_SERVER_URL=https://${{DEV_TUNNEL_URL}}";

  # Generate runtime settings to JSON file
  - uses: file/createOrUpdateJsonFile
    with:
{{#isNewProjectTypeEnabled}}
{{#PlaceProjectFileInSolutionDir}}
      target: ../local.settings.json
{{/PlaceProjectFileInSolutionDir}}
{{^PlaceProjectFileInSolutionDir}}
      target: ../{{appName}}/local.settings.json
{{/PlaceProjectFileInSolutionDir}}
{{/isNewProjectTypeEnabled}}
{{^isNewProjectTypeEnabled}}
      target: ./local.settings.json
{{/isNewProjectTypeEnabled}}
      content:
        IsEncrypted: false
        Values:
          FUNCTIONS_WORKER_RUNTIME: "dotnet-isolated"
          API_KEY: ${{SECRET_API_KEY}}

  # Register API KEY
  - uses: apiKey/register
    with:
      # Name of the API Key
      name: apiKey
      # Value of the API Key
      primaryClientSecret: ${{SECRET_API_KEY}}
      # Teams app ID
      appId: ${{TEAMS_APP_ID}}
      # Path to OpenAPI description document
      apiSpecPath: ./appPackage/apiSpecificationFile/repair.yml
    # Write the registration information of API Key into environment file for
    # the specified environment variable(s).
    writeToEnvironmentFile:
      registrationId: APIKEY_REGISTRATION_ID
  
  # Update API KEY
  - uses: apiKey/update
    with:
      # Name of the API Key
      name: apiKey      
      # Teams app ID
      appId: ${{TEAMS_APP_ID}}
      # Path to OpenAPI description document
      apiSpecPath: ./appPackage/apiSpecificationFile/repair.yml
      registrationId: ${{APIKEY_REGISTRATION_ID}}

  # Build Teams app package with latest env value
  - uses: teamsApp/zipAppPackage
    with:
      # Path to manifest template
      manifestPath: ./appPackage/manifest.json
      outputZipPath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip
      outputFolder: ./appPackage/build

  # Validate app package using validation rules
  - uses: teamsApp/validateAppPackage
    with:
      # Relative path to this file. This is the path for built zip file.
      appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip     

  # Apply the Teams app manifest to an existing Teams app in
  # Teams Developer Portal.
  # Will use the app id in manifest file to determine which Teams app to update.
  - uses: teamsApp/update
    with:
      # Relative path to this file. This is the path for built zip file.
      appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip

  # Extend your Teams app to Outlook and the Microsoft 365 app
  - uses: teamsApp/extendToM365
    with:
      # Relative path to the build app package.
      appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip
    # Write the information of created resources into environment file for
    # the specified environment variable(s).
    writeToEnvironmentFile:
      titleId: M365_TITLE_ID
      appId: M365_APP_ID
{{^isNewProjectTypeEnabled}}

  # Create or update debug profile in lauchsettings file
  - uses: file/createOrUpdateJsonFile
    with:
      target: ./Properties/launchSettings.json
      content:
        profiles:
        {{^DeclarativeCopilot}}
          Microsoft 365 app (browser):
            commandName: "Project"
            dotnetRunMessages: true
            launchBrowser: true
            launchUrl: "https://www.office.com/chat?auth=2"
            environmentVariables:
              ASPNETCORE_ENVIRONMENT: "Development"
            hotReloadProfile: "aspnetcore"
          Microsoft Teams (browser):
            commandName: "Project"
            commandLineArgs: "host start --port 5130 --pause-on-error"
            dotnetRunMessages: true
            launchBrowser: true
            launchUrl: "https://teams.microsoft.com?appTenantId=${{TEAMS_APP_TENANT_ID}}&login_hint=${{TEAMSFX_M365_USER_NAME}}"
            environmentVariables:
              ASPNETCORE_ENVIRONMENT: "Development"
            hotReloadProfile: "aspnetcore"
        {{/DeclarativeCopilot}}
        {{#DeclarativeCopilot}}
          "Copilot (browser)": {
            "commandName": "Project",
            "launchUrl": "https://m365.cloud.microsoft/chat/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429/${{AGENT_HINT}}?auth=2"
          }
        {{/DeclarativeCopilot}}
{{/isNewProjectTypeEnabled}}

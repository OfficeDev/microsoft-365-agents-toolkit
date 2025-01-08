# yaml-language-server: $schema=https://aka.ms/teams-toolkit/v1.8/yaml.schema.json
# Visit https://aka.ms/teamsfx-v5.0-guide for details on this file
# Visit https://aka.ms/teamsfx-actions for details on actions
version: v1.8

provision:
  # Creates a new Microsoft Entra app to authenticate users if
  # the environment variable that stores clientId is empty
  - uses: aadApp/create
    with:
      # Note: when you run aadApp/update, the Microsoft Entra app name will be updated
      # based on the definition in manifest. If you don't want to change the
      # name, make sure the name in Microsoft Entra manifest is the same with the name
      # defined here.
      name: {{appName}}
      # If the value is false, the action will not generate client secret for you
      generateClientSecret: false
      # Authenticate users with a Microsoft work or school account in your
      # organization's Microsoft Entra tenant (for example, single tenant).
      signInAudience: AzureADMyOrg
    # Write the information of created resources into environment file for the
    # specified environment variable(s).
    writeToEnvironmentFile:
      clientId: AAD_APP_CLIENT_ID
      objectId: AAD_APP_OBJECT_ID
      tenantId: AAD_APP_TENANT_ID
      authority: AAD_APP_OAUTH_AUTHORITY
      authorityHost: AAD_APP_OAUTH_AUTHORITY_HOST

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
        echo "::set-teamsfx-env OPENAPI_SERVER_DOMAIN=${{DEV_TUNNEL_URL}}";

  # Apply the Microsoft Entra manifest to an existing Microsoft Entra app. Will use the object id in
  # manifest file to determine which Microsoft Entra app to update.
  - uses: aadApp/update
    with:
      # Relative path to this file. Environment variables in manifest will
      # be replaced before apply to Microsoft Entra app
      manifestPath: ./aad.manifest.json
      outputFilePath: ./build/aad.manifest.${{TEAMSFX_ENV}}.json

  # Validate using manifest schema
  - uses: teamsApp/validateManifest
    with:
      # Path to manifest template
      manifestPath: ./appPackage/manifest.json

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

  # Generate runtime appsettings to JSON file
  - uses: file/createOrUpdateJsonFile
    with:
{{#isNewProjectTypeEnabled}}
{{#PlaceProjectFileInSolutionDir}}
      target: ../appsettings.Development.json
{{/PlaceProjectFileInSolutionDir}}
{{^PlaceProjectFileInSolutionDir}}
      target: ../{{appName}}/appsettings.Development.json
{{/PlaceProjectFileInSolutionDir}}
{{/isNewProjectTypeEnabled}}
{{^isNewProjectTypeEnabled}}
      target: ./appsettings.Development.json
{{/isNewProjectTypeEnabled}}
      content:
        CLIENT_ID: ${{AAD_APP_CLIENT_ID}}
        TENANT_ID: ${{AAD_APP_TENANT_ID}}
{{^isNewProjectTypeEnabled}}

  # Create or update debug profile in lauchsettings file
  - uses: file/createOrUpdateJsonFile
    with:
      target: ./Properties/launchSettings.json
      content:
        profiles:
          Microsoft Teams (browser):
            commandName: "Project"
            commandLineArgs: "host start --port 5130 --pause-on-error"
            dotnetRunMessages: true
            launchBrowser: true
            launchUrl: "https://teams.microsoft.com?appTenantId=${{TEAMS_APP_TENANT_ID}}&login_hint=${{TEAMSFX_M365_USER_NAME}}"
            environmentVariables:
              ASPNETCORE_ENVIRONMENT: "Development"
            hotReloadProfile: "aspnetcore"
{{/isNewProjectTypeEnabled}}
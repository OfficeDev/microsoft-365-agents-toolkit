# yaml-language-server: $schema=https://aka.ms/m365-agents-toolkits/v1.11/yaml.schema.json
# Visit https://aka.ms/teamsfx-v5.0-guide for details on this file
# Visit https://aka.ms/teamsfx-actions for details on actions
version: v1.11

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

  # Create AAD app for OAuth connection
  - uses: aadApp/create
    with:
      # The AAD app's display name
      name: {{appName}}-AAD${{APP_NAME_SUFFIX}}
      generateClientSecret: true
      generateServicePrincipal: true
      signInAudience: AzureADMyOrg
    writeToEnvironmentFile:
      clientId: AAD_APP_CLIENT_ID
      clientSecret: SECRET_AAD_APP_CLIENT_SECRET
      objectId: AAD_APP_OBJECT_ID
      tenantId: AAD_APP_TENANT_ID

  # Update AAD app with required API permissions
  - uses: aadApp/update
    with:
      manifestPath: ./aad.manifest.json
      outputFilePath: ./build/aad.manifest.${{TEAMSFX_ENV}}.json

  # If you want to enable SSO to integrate Microsoft 365 Retrieval API, uncomment this step and replace botFramework/create and script steps below
  # Create bot service with OAuth connection
  # - uses: arm/deploy
  #   with:
  #     subscriptionId: ${{AZURE_SUBSCRIPTION_ID}}
  #     resourceGroupName: ${{AZURE_RESOURCE_GROUP_NAME}}
  #     templates:
  #       - path: ./infra/azure.bicep
  #         parameters: ./infra/azure.parameters.local.json
  #         deploymentName: Create-resources-for-bot-local
  #     bicepCliVersion: v0.9.1
  
  # Local debugging steps without enabling SSO
  - uses: botFramework/create
    with:
      botId: ${{AAD_APP_CLIENT_ID}}
      name: {{appName}}
      messagingEndpoint: ${{BOT_ENDPOINT}}/api/messages
      description: ""
      channels:
        - name: msteams
  - uses: script
    with:
      run:
        echo "::set-teamsfx-env BOT_ID=${{AAD_APP_CLIENT_ID}}";

  # Generate runtime appsettings to JSON file
  - uses: file/createOrUpdateJsonFile
    with:
{{#isNewProjectTypeEnabled}}
{{#PlaceProjectFileInSolutionDir}}
      target: ../appsettings.Development.json
{{/PlaceProjectFileInSolutionDir}}
{{^PlaceProjectFileInSolutionDir}}
      target: ../{{ProjectName}}/appsettings.Development.json
{{/PlaceProjectFileInSolutionDir}}
{{/isNewProjectTypeEnabled}}
{{^isNewProjectTypeEnabled}}
      target: ./appsettings.Development.json
{{/isNewProjectTypeEnabled}}
      content:
        TokenValidation:
          Audiences:
            ClientId: ${{AAD_APP_CLIENT_ID}}
        Connections:
          BotServiceConnection:
            Settings:
              AuthorityEndpoint: "https://login.microsoftonline.com/${{AAD_APP_TENANT_ID}}"
              ClientId: ${{AAD_APP_CLIENT_ID}}
              ClientSecret: ${{SECRET_AAD_APP_CLIENT_SECRET}}
              TenantId: ${{AAD_APP_TENANT_ID}}
{{#useOpenAI}}
        OpenAI:
          ApiKey: ${{SECRET_OPENAI_API_KEY}}
{{/useOpenAI}}
{{#useAzureOpenAI}}
        Azure:
          OpenAIApiKey: ${{SECRET_AZURE_OPENAI_API_KEY}}
          OpenAIEndpoint: ${{AZURE_OPENAI_ENDPOINT}}
          OpenAIDeploymentName: ${{AZURE_OPENAI_DEPLOYMENT_NAME}}
{{/useAzureOpenAI}}

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
  # Developer Portal.
  # Will use the app id in manifest file to determine which Teams app to update.
  - uses: teamsApp/update
    with:
      # Relative path to this file. This is the path for built zip file.
      appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip

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
          Microsoft Teams (browser):
            commandName: "Project"
            dotnetRunMessages: true
            launchBrowser: true
            launchUrl: "https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&appTenantId=${{TEAMS_APP_TENANT_ID}}&login_hint=${{TEAMSFX_M365_USER_NAME}}"
            applicationUrl: "http://localhost:5130"
            environmentVariables:
              ASPNETCORE_ENVIRONMENT: "Development"
            hotReloadProfile: "aspnetcore"
{{/isNewProjectTypeEnabled}}

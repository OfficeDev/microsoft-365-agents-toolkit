# yaml-language-server: $schema=https://aka.ms/teams-toolkit/1.0.0/yaml.schema.json
#
# The teamsapp.yml composes automation tasks for Teams Toolkit when running other environment configurations.
# This file is used when selecting the Provision, Deploy, or Publish menu items in the Teams Toolkit for Visual Studio Code window
# or with the TeamsFx CLI commands.
# i.e. `teamsfx provision --env {environment name}` or `teamsfx deploy --env {environment name}`.
#
# You can customize this file. Visit https://aka.ms/teamsfx-v5.0-guide for more info about Teams Toolkit project files.
version: 1.0.0

environmentFolderPath: ./env

# Defines what the `provision` lifecycle step does with Teams Toolkit.
# Runs with the Provision menu or CLI using `teamsfx provision --env {environment name}`.
provision:
  # Automates the creation of a Teams app registration and saves the App ID to an environment file.
  - uses: teamsApp/create
    with:
      name: {{appName}}${{APP_NAME_SUFFIX}}
    writeToEnvironmentFile:
      teamsAppId: TEAMS_APP_ID

  # Creates a new Microsoft Entra app to authenticate users if
  # the environment variable that stores clientId is empty
  - uses: aadApp/create
    with:
      # Note: when you run aadApp/update, the Microsoft Entra app name will be updated
      # based on the definition in manifest. If you don't want to change the
      # name, make sure the name in Microsoft Entra manifest is the same with the name
      # defined here.
      name: {{appName}}${{APP_NAME_SUFFIX}}
      # If the value is false, the driver will not generate client secret for you
      generateClientSecret: true
      # organization's Microsoft Entra tenant (for example, single tenant).
      signInAudience: AzureADMultipleOrgs
    # Write the information of created resources into environment file for the
    # specified environment variable(s).
    writeToEnvironmentFile:
      clientId: BOT_ID
      # Environment variable that starts with `SECRET_` will be stored to the
      # .env.{envName}.user environment file
      clientSecret: SECRET_BOT_PASSWORD
      objectId: AAD_APP_OBJECT_ID
      tenantId: AAD_APP_TENANT_ID
      authority: AAD_APP_OAUTH_AUTHORITY
      authorityHost: AAD_APP_OAUTH_AUTHORITY_HOST

  # Apply the Microsoft Entra manifest to an existing Microsoft Entra app. Will use the object id in
  # manifest file to determine which Microsoft Entra app to update.
  - uses: aadApp/update
    with:
      # Relative path to this file. Environment variables in manifest will
      # be replaced before apply to Microsoft Entra app
      manifestPath: ./aad.manifest.json
      outputFilePath: ./build/aad.manifest.${{TEAMSFX_ENV}}.json

  - uses: arm/deploy # Deploy given ARM templates parallelly.
    env:
      # an arbitrary name for the connection
      OAUTH_CONNECTION_NAME: graph
    with:
      # AZURE_SUBSCRIPTION_ID is a built-in environment variable,
      # if its value is empty, TeamsFx will prompt you to select a subscription.
      # Referencing other environment variables with empty values
      # will skip the subscription selection prompt.
      subscriptionId: ${{AZURE_SUBSCRIPTION_ID}}
      # AZURE_RESOURCE_GROUP_NAME is a built-in environment variable,
      # if its value is empty, TeamsFx will prompt you to select or create one
      # resource group.
      # Referencing other environment variables with empty values
      # will skip the resource group selection prompt.
      resourceGroupName: ${{AZURE_RESOURCE_GROUP_NAME}}
      templates:
        - path: ./infra/azure.bicep # Relative path to this file
          # Relative path to this yaml file.
          # Placeholders will be replaced with corresponding environment
          # variable before ARM deployment.
          parameters: ./infra/azure.parameters.json
          # Required when deploying ARM template
          deploymentName: Create-resources-for-sso-teams-ai-${{TEAMSFX_ENV}}
      bicepCliVersion: v0.9.1

  # Optional: Automates schema and error checking of the Teams app manifest and outputs the results in the console.
  - uses: teamsApp/validateManifest
    with:
      manifestPath: ./appPackage/manifest.json

  # Automates creating a final app package (.zip) by replacing any variables in the manifest.json file for the current environment.
  - uses: teamsApp/zipAppPackage
    with:
      manifestPath: ./appPackage/manifest.json
      outputZipPath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip
      outputJsonPath: ./appPackage/build/manifest.${{TEAMSFX_ENV}}.json

  # Optional: Automates an app package check for errors that would prevent the app from being published and reports any problems.
  - uses: teamsApp/validateAppPackage
    with:
      appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip

  # Automates updating the Teams app manifest in Teams Developer Portal using the App ID from the manifest file.
  # This action ensures that any manifest changes are reflected when launching the app again in Teams.
  - uses: teamsApp/update
    with:
      appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip

# Defines what the `deploy` lifecycle step does with Teams Toolkit.
# Runs with the Deploy menu or CLI using `teamsfx deploy --env {environment name}`.
deploy:
  # Install any dependencies and build the web app using NPM
  - uses: cli/runNpmCommand
    name: install dependencies
    with:
      args: install

  - uses: cli/runNpmCommand
    name: build app
    with:
      args: run build --if-present

  # Deploy to an Azure App Service using the zip file created in the provision step.
  - uses: azureAppService/zipDeploy
    with:
      artifactFolder: .
      ignoreFile: .webappignore
      # This example uses the env var thats generated by the arm/deploy action.
      # You can replace it with an existing Azure Resource ID or other
      # custom environment variable.
      resourceId: ${{BOT_AZURE_APP_SERVICE_RESOURCE_ID}}

# Defines what the `publish` lifecycle step does with Teams Toolkit.
# Runs with the Deploy menu or CLI using `teamsfx publish --env {environment name}`.
publish:
  # Optional: Automates schema and error checking of the Teams app manifest and outputs the results in the console.
  - uses: teamsApp/validateManifest
    with:
      manifestPath: ./appPackage/manifest.json

  # Automates creating a final app package (.zip) by replacing any variables in the manifest.json file for the current environment.
  - uses: teamsApp/zipAppPackage
    with:
      manifestPath: ./appPackage/manifest.json
      outputZipPath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip
      outputJsonPath: ./appPackage/build/manifest.${{TEAMSFX_ENV}}.json

  # Optional: Automates an app package check for errors that would prevent the app from being published and reports any problems.
  - uses: teamsApp/validateAppPackage
    with:
      appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip

  # Automates updating the Teams app manifest in Teams Developer Portal using the App ID from the manifest file.
  # This action ensures that any manifest changes are reflected when launching the app again in Teams.
  - uses: teamsApp/update
    with:
      appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip

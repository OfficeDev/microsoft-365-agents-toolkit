# yaml-language-server: $schema=https://aka.ms/teams-toolkit/v1.8/yaml.schema.json
# Visit https://aka.ms/teamsfx-v5.0-guide for details on this file
# Visit https://aka.ms/teamsfx-actions for details on actions
version: v1.8

environmentFolderPath: ./env

# Triggered when 'teamsapp provision' is executed
provision:
  - uses: aadApp/create
    with:
      name: {{appName}}-${{TEAMSFX_ENV}}
      generateClientSecret: true
      signInAudience: AzureADMultipleOrgs
    writeToEnvironmentFile:
      clientId: AAD_APP_CLIENT_ID
      clientSecret: SECRET_AAD_APP_CLIENT_SECRET
      objectId: AAD_APP_OBJECT_ID
      tenantId: AAD_APP_TENANT_ID
      authority: AAD_APP_OAUTH_AUTHORITY
      authorityHost: AAD_APP_OAUTH_AUTHORITY_HOST

  - uses: aadApp/update
    with:
      manifestPath: ./aad.manifest.json
      outputFilePath: ./appPackage/build/aad.manifest.${{TEAMSFX_ENV}}.json

  - uses: file/createOrUpdateJsonFile
    with:
      target: ./local.settings.json
      appsettings:
        IsEncrypted: false
        Values:
          FUNCTIONS_WORKER_RUNTIME: node
          AzureWebJobsStorage: UseDevelopmentStorage=true
          AzureWebJobsFeatureFlags: EnableWorkerIndexing
          TEAMSFX_ENV: ${{TEAMSFX_ENV}}
          AZURE_CLIENT_ID: ${{AAD_APP_CLIENT_ID}}
          AZURE_CLIENT_SECRET: ${{SECRET_AAD_APP_CLIENT_SECRET}}
          AZURE_TENANT_ID: ${{AAD_APP_TENANT_ID}}
          CONNECTOR_ID: ${{CONNECTOR_ID}}
          CONNECTOR_NAME: ${{CONNECTOR_NAME}}
          CONNECTOR_DESCRIPTION: ${{CONNECTOR_DESCRIPTION}}
          CONNECTOR_REPOS: ${{CONNECTOR_REPOS}}
          CONNECTOR_ACCESS_TOKEN: ${{SECRET_CONNECTOR_ACCESS_TOKEN}}

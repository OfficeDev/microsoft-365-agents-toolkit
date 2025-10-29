# yaml-language-server: $schema=https://aka.ms/m365-agents-toolkits/v1.11/yaml.schema.json
version: v1.11

provision:
  # Create or reuse an existing Microsoft Entra application for bot.
  - uses: aadApp/create
    with:
      # The Microsoft Entra application's display name
      name: {{appName}}${{APP_NAME_SUFFIX}}
      generateClientSecret: true
      generateServicePrincipal: true
      signInAudience: AzureADMultipleOrgs
    writeToEnvironmentFile:
      # The Microsoft Entra application's client id created for bot.
      clientId: BOT_ID
      # The Microsoft Entra application's client secret created for bot.
      clientSecret: SECRET_BOT_PASSWORD
      # The Microsoft Entra application's object id created for bot.
      objectId: BOT_OBJECT_ID

  # Create or update the bot registration on dev.botframework.com
  - uses: botFramework/create
    with:
      botId: ${{BOT_ID}}
      name: {{appName}}
      messagingEndpoint: ${{APP_ENDPOINT}}/api/messages
      description: ""
      channels:
        - name: msteams

deploy:
  # Generate runtime environment variables for bot
  - uses: file/createOrUpdateEnvironmentFile
    with:
      target: ./.localConfigs
      envs:
        CLIENT_ID: ${{BOT_ID}}
        CLIENT_SECRET: ${{SECRET_BOT_PASSWORD}}
        TENANT_ID: ${{TEAMS_APP_TENANT_ID}}
        {{#useOpenAI}}
        OPENAI_API_KEY: ${{SECRET_OPENAI_API_KEY}}
        {{/useOpenAI}}
        {{#useAzureOpenAI}}
        AZURE_OPENAI_API_KEY: ${{SECRET_AZURE_OPENAI_API_KEY}}
        AZURE_OPENAI_ENDPOINT: ${{AZURE_OPENAI_ENDPOINT}}
        AZURE_OPENAI_DEPLOYMENT_NAME: ${{AZURE_OPENAI_DEPLOYMENT_NAME}}
        AZURE_OPENAI_EMBEDDING_DEPLOYMENT_NAME: ${{AZURE_OPENAI_EMBEDDING_DEPLOYMENT_NAME}}
        {{/useAzureOpenAI}}
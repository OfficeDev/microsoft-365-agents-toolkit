# yaml-language-server: $schema=https://aka.ms/m365-agents-toolkits/v1.9/yaml.schema.json
# Visit https://aka.ms/teamsfx-v5.0-guide for details on this file
# Visit https://aka.ms/teamsfx-actions for details on actions
version: v1.9

deploy:
  # Install development tool(s)
  - uses: devTool/install
    with:
      testTool:
        version: ~0.2.1
        symlinkDir: ./devTools/teamsapptester
      nodejs:
        symlinkDir: ./devTools/nodejs

  # Generate runtime environment variables
  - uses: file/createOrUpdateEnvironmentFile
    with:
      target: ./.env
      envs:
        TEAMSFX_NOTIFICATION_STORE_FILENAME: ${{TEAMSFX_NOTIFICATION_STORE_FILENAME}}
        BOT_ID: ""
        BOT_PASSWORD: ""
        {{#useOpenAI}}
        OPENAI_API_KEY: ${{SECRET_OPENAI_API_KEY}}
        {{/useOpenAI}}
        {{#useAzureOpenAI}}
        AZURE_OPENAI_API_KEY: ${{SECRET_AZURE_OPENAI_API_KEY}}
        AZURE_OPENAI_MODEL_DEPLOYMENT_NAME: ${{AZURE_OPENAI_MODEL_DEPLOYMENT_NAME}}
        AZURE_OPENAI_ENDPOINT: ${{AZURE_OPENAI_ENDPOINT}}
        {{/useAzureOpenAI}}
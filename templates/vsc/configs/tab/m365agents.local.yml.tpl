# yaml-language-server: $schema=https://aka.ms/m365-agents-toolkits/v1.9/yaml.schema.json
# Visit https://aka.ms/teamsfx-v5.0-guide for details on this file
# Visit https://aka.ms/teamsfx-actions for details on actions
version: v1.9

provision:
  # Set APP_DOMAIN and APP_ENDPOINT for local launch
  - uses: script
    with:
      run:
        echo "::set-teamsfx-env APP_DOMAIN=localhost";
        echo "::set-teamsfx-env APP_ENDPOINT=https://localhost:3978";

deploy:
  # Install development tool(s)
  - uses: devTool/install
    with:
      devCert:
        trust: true
    # Write the information of installed development tool(s) into environment
    # file for the specified environment variable(s).
    writeToEnvironmentFile:
      sslCertFile: SSL_CRT_FILE
      sslKeyFile: SSL_KEY_FILE

  # Generate runtime environment variables
  - uses: file/createOrUpdateEnvironmentFile
    with:
      target: ./.localConfigs
      envs:
        SSL_CRT_FILE: ${{SSL_CRT_FILE}}
        SSL_KEY_FILE: ${{SSL_KEY_FILE}}

# TeamsFx-UI-Test

A UI Test Project based on [Vscode Extension Tester](https://github.com/redhat-developer/vscode-extension-tester/wiki)

## Getting started

### Prerequisites

- node >= 18
- NPM >= 8
- m365 account
- azure account

### Setup

- (**Required**) Run `npm install @microsoft/m365agentstoolkit-cli@alpha` to download latest CLI
- (**Options**) If test migration, run `npm install @microsoft/teamsfx-cli@1.2.6` to download old version CLI
- (**Required**) Run `pnpm install`
- (**Required**) Run `npm run build`
- (**Required**) Login your m365 account via TeamsFx extension
- (**Required**) Login your azure account via TeamsFx extension
- (**Required**) Setup local environment variables. Create .env file and add following variables (these variables are for verifying installed Teams App, so the test needs to login on-behalf-of you):

```
CLEAN_TENANT_ID=
CLEAN_CLIENT_ID=

AZURE_ACCOUNT_NAME=
AZURE_ACCOUNT_PASSWORD=
AZURE_SUBSCRIPTION_ID=
AZURE_SUBSCRIPTION_NAME=
AZURE_TENANT_ID=

M365_ACCOUNT_NAME=
M365_ACCOUNT_PASSWORD=
M365_DISPLAY_NAME=
M365_COLLABORATOR=

TEAMSFX_DEV_TUNNEL_TEST=true
TEAMSFX_V3=true
TEAMSFX_V3_MIGRATION=true
TEAMSFX_TELEMETRY_TEST=true
TARGET_CLI_VERSION=

CI_ENABLED=true
```

- (**Required**) Run `npx extest get-vscode --storage .test-resources --type stable` to download vscode
- (**Required**) Run `npx extest get-chromedriver --storage .test-resources --type stable` to download chromedriver
- (**Required**) Download TeamsFx vsix file to this project root folder. You can download it from the [artifacts of TeamsFx CD action](https://github.com/OfficeDev/TeamsFx/actions/workflows/cd.yml). Remember to unzip.
- (**Required**) Run `npx extest install-vsix --storage .test-resources --extensions_dir .test-resources --type stable --vsix_file ${{ YOUR VSIX FILE NAME }} ` to install Teams Toolkit
- (**OPTIONAL**) If local test docker cases, Run `npx extest install-from-marketplace --storage .test-resources --extensions_dir .test-resources --type stable ms-azuretools.vscode-docker` to install docker extension.
- (**Required**) Run `npx extest run-tests --storage .test-resources --extensions_dir .test-resources --type stable --code_settings ./settings.json ./out/ui-test/**/${{ YOUR TEST CASE }}.test.js` to execute your case
- (**OPTIONAL**) If you want to debug your case via vscode, replace "YOUR TEST CASE" with your case name in .vscode/launch.json and click F5

### How to add a new test case

The new test case can be added in the directory `./src/ui-test`. The test can inherit the base class `TestContext`.

There are some common VSCode operations in `./src/vscodeOperation.ts`. [Here](https://github.com/redhat-developer/vscode-extension-tester) you can learn more about vscode-extension-tester.

If your test case needs to open the browser, sideloading an Teams App and verify the Team App, you can find some common Teams App operations in `./src/playwrightOperation.ts`. [Here](https://playwright.dev/docs/intro) you can learn more about playwright.

If you want to add your case to schedualed job, you can update the `.github/workflows/pvt.json`, choose target os, node version and add your file name to it.
For example, if your test file is `src/ui-test/localdebug/localdebug-bot-ts.test.ts`, and you want to execute it on windows with node 18, then add `localdebug-bot-ts` to

```json
{
  "windows-latest": {
    "node-18": ["localdebug-bot-ts"]
  }
}
```

### How to execute for WSL2 users

- (In all the steps, use the non-Windows version scripts)
- Install an X Server (e.g. VcXsrv), start it on display N (e.g. display 1) and disable authentication.
- Install the dependencies of vscode in WSL2. Refer to [this document](https://code.visualstudio.com/docs/setup/linux#_debian-and-ubuntu-based-distributions).
- Run the test: use either way, vscode or manually:
  - vscode:
    - Run "Debug UI Test WSL2"
  - Manual way
    - Set `DISPLAY` environment variable to HostIP:DISPLAY_ID (e.g. 172.16.0.1:1, display number from last step).
    - Set `DONT_PROMPT_WSL_INSTALL` environment variable to a non empty string. This is to ensure vscode in WSL does not use the Windows version.
- You can also refer to [this tutorial](https://www.gregbrisebois.com/posts/chromedriver-in-wsl2/)
- If you need to use Azure account (i.e. use Azure Account extension), you need to [configure your keyring manually](https://github.com/atom/node-keytar/issues/132#issuecomment-444159414). Otherwise, you may encounter the `Unknown or unsupported transport 'disabled' for address 'disabled:'` error.

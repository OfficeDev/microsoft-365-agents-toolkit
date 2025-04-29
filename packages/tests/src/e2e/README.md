# CLI testing flow 

## Creating your test case

### Preparation

1. Install Teamsfx-cli 

   ```bash
   # download test version
   npm install -g @microsoft/m365agentstoolkit-cli@VERSION
   ```

2. Download Teamsfx Project

   ```bash
   git clone https://github.com/OfficeDev/Teamsfx.git
   ```

3. Create your own branch 

   ```bash
   # In Teamsfx Folder
   git checkout -b <your_branch>
   ```

4. CLI testing case should  in folder "packages/cli/tests/e2e/"

   ![cli_test_file_path](../../img/cli_test_file_path.png)

### Create

Use mocha framework to build your test

#### **fileName**

Naming notations must follow below:

**normal test: XXX.tests.ts**

**v3 test: XXX.v3.tests.ts**

**dotnet test: XXX.dotnet.tests.ts**

#### **structure**

```typescript
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author your name <your_email@microsoft.com> 
 */
import { it } from "@microsoft/extra-shot-mocha";
describe("test name", function () {
    before(() => {
			// do something setup
  });

  after(async () => {
		// do something cleanup
  });
  
  // we suggest one file has only one 'it'
  it(`step name`, { testPlanCaseId: XXXXXX }, async function () {
		// do testing bellow here
    
  });
});

```

#### **Command**

- CliHelper.ts file contain teamsfx-cli command, we suggest to use it to excute CLI commands.   

- constants.ts file contain all commands constant. Like 'teamsfx new, teamsfx provision...'.

```typescript
import { CliHelper } from "../../commonlib/cliHelper";
import { Capability } from "../../commonlib/constants";
```

### validate

- In 'tests/commonlib/' folder, We provide some validator to varify the automation.

- Also you can use 'expect' validate from chai.

```typescript
import { expect } from "chai";
```

### clean

Always cleanup when finish auto test. You can use 'cleanUp' function to clean your resourse.

```typescript
import { cleanUp } from "../commonUtils";

// ...
  after(async () => {
		// demo
    await cleanUp(appName, projectPath, true, false, false);
  });

// ...

```

### test

**Github Action**

Use GitHub Action to automate your test 

![github_action_e2e](../../img/github_action_e2e.png)

To trigger the runner, you need to manually click "Run workflow".

select your own branch. (Make sure your change already 'git push' to your branch) and follow the instruction to input your test file path.

![run_workflow](../../img/run_workflow.png)

## Example

example file path: packages/cli/tests/e2e/template/ProvisionHelloWorldBot.tests.ts

```typescript
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { expect } from "chai";
import fs from "fs-extra";
import path from "path";
import { it } from "@microsoft/extra-shot-mocha";
import {
  execAsync,
  getTestFolder,
  cleanUp,
  setSimpleAuthSkuNameToB1Bicep,
  getSubscriptionId,
  readContextMultiEnv,
  getUniqueAppName,
} from "../commonUtils";
import { AadValidator, BotValidator } from "../../commonlib";
import { TemplateProject } from "../../commonlib/constants";
import { CliHelper } from "../../commonlib/cliHelper";
import m365Login from "@microsoft/m365agentstoolkit-cli/src/commonlib/m365Login";
import { environmentNameManager } from "@microsoft/teamsfx-core";

describe("teamsfx new sample", function () {
  const testFolder = getTestFolder();
  const subscription = getSubscriptionId();
  const appName = getUniqueAppName();
  const projectPath = path.resolve(testFolder, appName);
  const env = environmentNameManager.getDefaultEnvName();

  it(`${TemplateProject.HelloWorldBotSSO}`, { testPlanCaseId: 15277464 }, async function () {
    await CliHelper.createTemplateProject(
      appName,
      testFolder,
      TemplateProject.HelloWorldBotSSO,
      TemplateProject.HelloWorldBotSSO
    );

    expect(fs.pathExistsSync(projectPath)).to.be.true;
    expect(fs.pathExistsSync(path.resolve(projectPath, ".fx"))).to.be.true;

    // Provision
    await setSimpleAuthSkuNameToB1Bicep(projectPath, env);
    await CliHelper.provisionProject(projectPath);

    // Validate Provision
    const context = await readContextMultiEnv(projectPath, env);

    // Validate Bot Provision
    const bot = new BotValidator(context, projectPath, env);
    await bot.validateProvision(false);

    // deploy
    await CliHelper.deployAll(projectPath);

    {
      // Validate deployment

      // Get context
      const context = await fs.readJSON(`${projectPath}/.fx/states/state.dev.json`);

      // Validate Aad App
      const aad = AadValidator.init(context, false, m365Login);
      await AadValidator.validate(aad);

      // Validate Bot Deploy
      const bot = new BotValidator(context, projectPath, env);
      await bot.validateDeploy();
    }

    // test (validate)
    await execAsync(`teamsfx validate`, {
      cwd: projectPath,
      env: process.env,
      timeout: 0,
    });

    // package
    await execAsync(`teamsfx package`, {
      cwd: projectPath,
      env: process.env,
      timeout: 0,
    });
  });

  after(async () => {
    await cleanUp(appName, projectPath, true, true, false);
  });
});
```

## Methods

### CliHelper


#### addEnv

```bash
teamsfx env add ${env} --env dev
```

#### provisionProject

```bash
teamsfx provision ${option}
```

#### addApiConnection

authType: ['cert', 'apiKey', 'basic', 'custom', 'aad']

```bash
teamsfx add api-connection ${authType} ${commonInputs} ${options} --interactive false
```

#### addCICDWorkflows

```bash
teamsfx add cicd ${option}
```

#### addExistingApi

```bash
teamsfx add api-connection ${option}
```

#### deployAll

```bash
teamsfx deploy ${option}
```

#### deployProject

```typescript
enum ResourceToDeploy {
  Spfx = "spfx",
  FrontendHosting = "frontend-hosting",
  Bot = "bot",
  Function = "function",
  Apim = "apim",
  AadManifest = "aad-manifest",
}
```

```bash
teamsfx deploy ${resourceToDeploy} ${option}
```

#### createDotNetProject

capability: "tab" | "bot"

```bash
teamsfx new --interactive false --runtime dotnet --app-name ${appName} --capability ${capability} ${options}
```

#### createProjectWithCapability

```typescript
enum Capability {
  Tab = "tab",
  SSOTab = "sso-tab",
  Bot = "bot",
  MessageExtension = "message-extension",
  M365SsoLaunchPage = "sso-launch-page",
  M365SearchApp = "search-app",
  ExistingTab = "existing-tab",
  TabSso = "TabSSO",
  BotSso = "BotSSO",
  TabNonSso = "tab-non-sso",
  Notification = "notification",
}
```

```bash
teamsfx new --interactive false --app-name ${appName} --capability ${capability} ${options}
```

#### createTemplateProject

```typescript
enum TemplateProject {
  HelloWorldTabSSO = "hello-world-tab",
  HelloWorldTabBackEnd = "hello-world-tab-with-backend",
  HelloWorldBot = "hello-world-bot",
  ContactExporter = "graph-toolkit-contact-exporter",
  OneProductivityHub = "graph-toolkit-one-productivity-hub",
  HelloWorldBotSSO = "bot-sso",
  TodoListBackend = "todo-list-with-Azure-backend",
  TodoListSpfx = "todo-list-spfx",
  ShareNow = "share-now",
  MyFirstMeeting = "hello-world-in-meeting",
  queryOrg = "query-org-user-with-message-extension-sso",
  TodoListM365 = "todo-list-with-azure-backend-m365",
  NpmSearch = "npm-search-connector-m365",
  HelloWorldTab = "hello-world-tab-without-sso",
  ProactiveMessaging = "bot-proactive-messaging-teamsfx",
  AdaptiveCard = "adaptive-card-notification",
  IncomingWebhook = "incoming-webhook-notification",
  GraphConnector = "graph-connector-app",
  StockUpdate = "stocks-update-notification-bot",
}
```

```bash
teamsfx new sample ${template} --interactive false 
```

#### addCapabilityToProject

```bash
teamsfx add ${capabilityToAdd}
```

#### addResourceToProject

```typescript
enum Resource {
  AzureKeyVault = "azure-keyvault",
  AzureFunction = "azure-function",
  AzureApim = "azure-apim",
  AzureSql = "azure-sql",
}
```

```bash
teamsfx add ${resourceToAdd} ${options}
```

#### getUserSettings

```bash
teamsfx config get ${key} --env ${env}
```

### Validator Class

#### aadValidate

#### apimValidator

#### appStudioValidator

#### botValidator

#### existingAppValidator

#### frontendValidator

#### functionValidator

#### keyVaultValidator

#### m365Validator

#### sharepointValidator

#### simpleAuthValidator

#### sqlvalidator

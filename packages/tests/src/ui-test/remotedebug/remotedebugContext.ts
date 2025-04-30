// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs-extra";
import * as path from "path";
import { TestContext } from "../testContext";
import {
  CommandPaletteCommands,
  Timeout,
  TestFilePath,
  Notification,
} from "../../utils/constants";
import { Env } from "../../utils/env";
import {
  cleanUpAadApp,
  cleanTeamsApp,
  cleanAppStudio,
  cleanUpLocalProject,
  cleanUpResourceGroup,
  createResourceGroup,
} from "../../utils/cleanHelper";
import {
  execCommandIfExist,
  getNotification,
  clearNotifications,
} from "../../utils/vscodeOperation";
import { ModalDialog, InputBox, VSBrowser } from "vscode-extension-tester";
import { dotenvUtil } from "../../utils/envUtil";
import { execAsync } from "../../utils/commonUtils";
import { CliHelper } from "../cliHelper";

export class RemoteDebugTestContext extends TestContext {
  public testName: string;

  constructor(testName: string) {
    super(testName);
    this.testName = testName;
  }

  async getTeamsAppId(projectPath: string, envName = "dev"): Promise<string> {
    const userDataFile = path.join(
      TestFilePath.configurationFolder,
      `.env.${envName}`
    );
    const configFilePath = path.resolve(projectPath, userDataFile);
    const context = dotenvUtil.deserialize(
      await fs.readFile(configFilePath, { encoding: "utf8" })
    );
    const result = context.obj.TEAMS_APP_ID as string;
    console.log(`TEAMS APP ID: ${result}`);
    return result;
  }

  public async after(hasAadPlugin = true, hasBotPlugin = false) {
    await this.context!.close();
    await this.browser!.close();
  }

  public async cleanUp(
    appName: string,
    projectPath: string,
    hasAadPlugin = true,
    hasBotPlugin = false,
    hasApimPlugin = false,
    envName = "dev"
  ) {
    const cleanUpAadAppPromise = cleanUpAadApp(
      projectPath,
      hasAadPlugin,
      hasBotPlugin,
      hasApimPlugin,
      envName
    );
    return Promise.all([
      // delete aad app
      cleanUpAadAppPromise,
      // uninstall Teams app
      cleanTeamsApp(appName),
      // delete Teams app in app studio
      cleanAppStudio(appName),
      // remove resouce group
      cleanUpResourceGroup(appName, envName),
      // remove project
      cleanUpLocalProject(projectPath, cleanUpAadAppPromise),
    ]);
  }
}

export async function getAadObjectId(
  projectPath: string,
  envName = "dev"
): Promise<string> {
  const userDataFile = path.join(
    TestFilePath.configurationFolder,
    `.env.${envName}`
  );
  const configFilePath = path.resolve(projectPath, userDataFile);
  const context = dotenvUtil.deserialize(
    await fs.readFile(configFilePath, { encoding: "utf8" })
  );
  const result = context.obj.AAD_APP_OBJECT_ID as string;
  console.log(`TEAMS APP OBJECT ID: ${result}`);
  return result;
}

export async function setSimpleAuthSkuNameToB1(projectPath: string) {
  const simpleAuthPluginName = "fx-resource-simple-auth";
  const envFilePathSuffix = path.join(".fx", "env.default.json");
  const envFilePath = path.resolve(projectPath, envFilePathSuffix);
  const context = await fs.readJSON(envFilePath);
  context[simpleAuthPluginName]["skuName"] = "B1";
  return fs.writeJSON(envFilePath, context, { spaces: 4 });
}

export async function setSkuNameToB1(projectPath: string) {
  const parameters = "parameters";
  const webAppSku = "webAppSku";
  const azureParametersFilePathSuffix = path.join(
    "infra",
    "azure.parameters.json"
  );
  const azureParametersFilePath = path.resolve(
    projectPath,
    azureParametersFilePathSuffix
  );
  const context = await fs.readJSON(azureParametersFilePath);
  try {
    context[parameters][webAppSku]["value"] = "B1";
  } catch {
    console.log("Cannot set the propertie.");
  }
  return fs.writeJSON(azureParametersFilePath, context, { spaces: 4 });
}

export async function setSkuNameToStandard(projectPath: string) {
  const parameters = "parameters";
  const staticWebAppSku = "staticWebAppSku";
  const azureParametersFilePathSuffix = path.join(
    "infra",
    "azure.parameters.json"
  );
  const azureParametersFilePath = path.resolve(
    projectPath,
    azureParametersFilePathSuffix
  );
  const context = await fs.readJSON(azureParametersFilePath);
  try {
    context[parameters][staticWebAppSku]["value"] = "Standard";
  } catch {
    console.log("Cannot set the propertie.");
  }
  return fs.writeJSON(azureParametersFilePath, context, { spaces: 4 });
}

export async function setSimpleAuthSkuNameToB1Bicep(
  projectPath: string,
  envName: string
) {
  const ConfigFolderName = "fx";
  const InputConfigsFolderName = "configs";
  const bicepParameterFile = path.join(
    `.${ConfigFolderName}`,
    InputConfigsFolderName,
    `azure.parameters.${envName}.json`
  );
  const parametersFilePath = path.resolve(projectPath, bicepParameterFile);
  const parameters = await fs.readJSON(parametersFilePath);
  parameters["parameters"]["provisionParameters"]["value"]["simpleAuthSku"] =
    "B1";
  return fs.writeJSON(parametersFilePath, parameters, { spaces: 4 });
}

export async function setBotSkuNameToB1Bicep(
  projectPath: string,
  filePath = ""
) {
  const azureParametersFilePathSuffix = filePath
    ? path.join(filePath)
    : path.join("infra", "azure.parameters.json");
  const azureParametersFilePath = path.resolve(
    projectPath,
    azureParametersFilePathSuffix
  );
  const ProvisionParameters = await fs.readJSON(azureParametersFilePath);
  ProvisionParameters["parameters"]["provisionParameters"]["value"][
    "botWebAppSKU"
  ] = "B1";
  return fs.writeJSON(azureParametersFilePath, ProvisionParameters, {
    spaces: 4,
  });
}

export async function inputSqlUserName(
  input: InputBox,
  sqlAdminName: string,
  sqlPassword: string
) {
  await input.setText(sqlAdminName);
  await input.confirm();
  await input.setText(sqlPassword);
  await input.confirm();
  await input.setText(sqlPassword);
  await input.confirm();
}

export async function provisionProject(
  appName: string,
  projectPath = "",
  createRg = true,
  tool: "ttk" | "cli" = "cli",
  option = "",
  env: "dev" | "local" = "dev",
  processEnv?: NodeJS.ProcessEnv
) {
  if (tool === "cli") {
    await runCliProvision(
      projectPath,
      appName,
      createRg,
      option,
      env,
      processEnv
    );
  } else {
    await runProvision(appName);
  }
}

export async function deployProject(
  projectPath: string,
  waitTime: number = Timeout.tabDeploy,
  tool: "ttk" | "cli" = "cli",
  option = "",
  env: "dev" | "local" = "dev",
  processEnv?: NodeJS.ProcessEnv,
  retries?: number,
  newCommand?: string
) {
  if (tool === "cli") {
    await runCliDeploy(
      projectPath,
      option,
      env,
      processEnv,
      retries,
      newCommand
    );
  } else {
    await runDeploy(waitTime);
  }
}

export async function runCliProvision(
  projectPath: string,
  appName: string,
  createRg = true,
  option = "",
  env: "dev" | "local" = "dev",
  processEnv?: NodeJS.ProcessEnv
) {
  if (createRg) {
    await createResourceGroup(appName, env, "westus");
  }
  const resourceGroupName = `${appName}-${env}-rg`;
  await CliHelper.showVersion(projectPath, processEnv);
  await CliHelper.provisionProject2(projectPath, option, env, {
    ...process.env,
    AZURE_RESOURCE_GROUP_NAME: resourceGroupName,
  });
}

export async function runCliDeploy(
  projectPath: string,
  option = "",
  env: "dev" | "local" = "dev",
  processEnv?: NodeJS.ProcessEnv,
  retries?: number,
  newCommand?: string
) {
  await CliHelper.deployAll(
    projectPath,
    option,
    env,
    processEnv,
    retries,
    newCommand
  );
}

export async function runProvision(
  appName: string,
  envName = "dev",
  containSql = false,
  spfx = false
) {
  console.log("start to provision");
  await execCommandIfExist(CommandPaletteCommands.ProvisionCommand);
  const driver = VSBrowser.instance.driver;
  await driver.sleep(Timeout.shortTimeWait);

  if (!spfx) {
    const subscriptionConfirmInput = await InputBox.create();
    const subscriptionConfirmInputTitle =
      await subscriptionConfirmInput.getTitle();
    if (subscriptionConfirmInputTitle?.includes("Select a Subscription")) {
      await subscriptionConfirmInput.selectQuickPick(Env.azureSubscriptionName);
      await driver.sleep(Timeout.shortTimeWait);
      console.log(
        `The subscription ${Env.azureSubscriptionName} is setting up.`
      );
    }

    const provisionConfirmInput = await InputBox.create();
    await provisionConfirmInput.selectQuickPick("+ New resource group");
    await driver.sleep(Timeout.shortTimeWait);
    const rgName = `${appName}-${envName}-rg`;
    console.log("new resource group: ", rgName);
    await provisionConfirmInput.setText(rgName);
    await provisionConfirmInput.confirm();
    await driver.sleep(Timeout.shortTimeWait);
    // await provisionConfirmInput.selectQuickPick("East US");
    await provisionConfirmInput.setText("West US");
    await provisionConfirmInput.confirm();
    console.log("location: West US");
    await driver.sleep(Timeout.shortTimeWait);
    const dialog = new ModalDialog();
    console.log("click provision button");
    await dialog.pushButton("Provision");
    await driver.sleep(Timeout.shortTimeLoading);

    const waitTime = Timeout.tabProvision;
    if (containSql) {
      const provisionSqlUserInput = await InputBox.create();
      await inputSqlUserName(provisionSqlUserInput, "Abc123321", "Cab232332");
    }
    await driver.sleep(waitTime);
  }

  try {
    await getNotification(
      Notification.ProvisionSucceeded,
      Timeout.longTimeWait,
      8,
      ["Error", "Failed"]
    );
    console.log("provision successfully");
    return;
  } catch {
    await clearNotifications();
    await reRunProvision();
  }
  await getNotification(
    Notification.ProvisionSucceeded,
    Timeout.longTimeWait,
    8,
    ["Error", "Failed"]
  );
  console.log("provision successfully");
}

export async function reRunProvision() {
  console.log("start to rerun provision");
  await execCommandIfExist(CommandPaletteCommands.ProvisionCommand);
  const driver = VSBrowser.instance.driver;
  await driver.sleep(Timeout.shortTimeWait);
  const dialog = new ModalDialog();
  console.log("click provision button");
  await dialog.pushButton("Provision");
  await driver.sleep(Timeout.shortTimeLoading);
  const waitTime = Timeout.tabProvision;
  await driver.sleep(waitTime);
}

export async function runDeploy(waitTime: number = Timeout.tabDeploy) {
  const driver = VSBrowser.instance.driver;
  await clearNotifications();
  console.log("start to deploy");
  await execCommandIfExist(CommandPaletteCommands.DeployCommand);

  await driver.sleep(Timeout.shortTimeWait);
  const dialog = new ModalDialog();
  console.log("click deploy button");
  await dialog.pushButton("Deploy");
  await driver.sleep(waitTime);

  try {
    await getNotification(
      Notification.DeploySucceeded,
      Timeout.longTimeWait,
      8,
      ["Error", "Failed"]
    );
    console.log("deploy successfully");
    return;
  } catch {
    await clearNotifications();
    await reRunDeploy(waitTime);
  }
  await getNotification(Notification.DeploySucceeded, Timeout.longTimeWait, 8, [
    "Error",
    "Failed",
  ]);
  console.log("deploy successfully");
}

export async function reRunDeploy(waitTime: number = Timeout.tabDeploy) {
  const driver = VSBrowser.instance.driver;
  console.log("start to rerun deploy");
  await execCommandIfExist(CommandPaletteCommands.DeployCommand);

  await driver.sleep(Timeout.shortTimeWait);
  const dialog = new ModalDialog();
  console.log("click deploy button");
  await dialog.pushButton("Deploy");

  await driver.sleep(waitTime);
}

export async function runPublish(rePublish = false) {
  await clearNotifications();
  await execCommandIfExist(CommandPaletteCommands.PublishCommand);
  const driver = VSBrowser.instance.driver;

  // const confirmInput = await InputBox.create();
  // await confirmInput.selectQuickPick("Install for your organization");

  if (rePublish) {
    await driver.sleep(Timeout.longTimeWait);
    const dialog = new ModalDialog();
    await dialog.pushButton("Confirm");
  }

  await driver.sleep(Timeout.longTimeWait);

  await getNotification(Notification.PublishSucceeded, Timeout.longTimeWait);
}

export async function selectEnv(envName = "dev") {
  const driver = VSBrowser.instance.driver;
  const input = await InputBox.create();
  await driver.sleep(Timeout.input);
  await input.selectQuickPick(envName);
}

export async function createEnv(envName = "staging") {
  const driver = VSBrowser.instance.driver;
  await execCommandIfExist(
    CommandPaletteCommands.CreateEnvironmentCommand,
    Timeout.webView
  );
  const input = await InputBox.create();
  // Input App Name
  await input.setText(envName);
  await driver.sleep(Timeout.input);
  await input.confirm();
  await driver.sleep(Timeout.shortTimeWait);
  console.log(`Created env ${envName}`);
}

export async function setSkipAddingSqlUser(
  projectPath: string,
  envName = "dev"
) {
  const ConfigFolderName = "fx";
  const InputConfigsFolderName = "configs";
  const configParameterFile = path.join(
    `.${ConfigFolderName}`,
    InputConfigsFolderName,
    `config.${envName}.json`
  );
  const parametersFilePath = path.resolve(projectPath, configParameterFile);
  const parameters = await fs.readJSON(parametersFilePath);
  parameters["skipAddingSqlUser"] = true;
  return fs.writeJSON(parametersFilePath, parameters, { spaces: 4 });
}

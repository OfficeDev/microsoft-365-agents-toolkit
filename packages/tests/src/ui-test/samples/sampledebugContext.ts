// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs-extra";
import * as path from "path";
import {
  Timeout,
  TreeViewCommands,
  TemplateProject,
  TemplateProjectFolder,
  TestFilePath,
} from "../../utils/constants";
import { dotenvUtil } from "../../utils/envUtil";
import { InputBox, VSBrowser } from "vscode-extension-tester";
import { getSampleAppName } from "../../utils/nameUtil";
import {
  execCommandIfExistFromTreeView,
  openExistingProject,
  stopDebugging,
  clearNotifications,
} from "../../utils/vscodeOperation";
import { assert, expect } from "chai";
import { TestContext } from "../testContext";
import * as dotenv from "dotenv";
import { CliHelper } from "../cliHelper";
import { AzSqlHelper } from "../../utils/azureCliHelper";
import {
  cleanUpAadApp,
  cleanTeamsApp,
  cleanAppStudio,
  cleanUpLocalProject,
  cleanUpResourceGroup,
  createResourceGroup,
} from "../../utils/cleanHelper";
import { Executor } from "../../utils/executor";
import { runProvision, runDeploy } from "../remotedebug/remotedebugContext";

export class SampledebugContext extends TestContext {
  public readonly appName: string;
  public readonly sampleName: TemplateProject;
  public readonly projectPath: string;
  public originPath = "";
  public readonly testRootFolder: string;
  public env: "dev" | "local" = "dev";
  public originSample: TemplateProjectFolder;
  public rgName: string;
  public readonly repoPath: string;

  constructor(
    sampleName: TemplateProject,
    originSample: TemplateProjectFolder,
    testRootFolder = "./resource",
    repoPath = "./resource"
  ) {
    super(sampleName);
    this.sampleName = sampleName;
    this.originSample = originSample;
    this.repoPath = repoPath;
    if (sampleName.length >= 20) {
      this.appName = getSampleAppName(
        sampleName
          .split(" ")
          .splice(0, 3)
          .join("")
          .split(",")
          .join("")
          .split(")")
          .join("")
          .split("(")
          .join("")
      );
    } else {
      this.appName = getSampleAppName(sampleName);
    }
    // fix eslint error
    this.testRootFolder = testRootFolder;
    this.projectPath = path.resolve(this.testRootFolder, this.appName);
    this.env = "dev";
    this.rgName = `${this.appName}-dev-rg`;
  }

  public async sampleAfter(
    rgName: string,
    hasAadPlugin = true,
    hasBotPlugin = false,
    envName = "dev"
  ): Promise<void> {
    await stopDebugging();
    await this.context?.close();
    await this.browser?.close();
    await this.cleanUp(
      this.appName,
      this.projectPath,
      hasAadPlugin,
      hasBotPlugin,
      false,
      envName
    );
  }

  public async after(
    hasAadPlugin = true,
    hasBotPlugin = false,
    envName = "local"
  ): Promise<void> {
    await stopDebugging();
    await this.context?.close();
    await this.browser?.close();
    await this.cleanUp(
      this.appName,
      this.projectPath,
      hasAadPlugin,
      hasBotPlugin,
      false,
      envName
    );
  }

  public async cleanUp(
    appName: string,
    projectPath: string,
    hasAadPlugin = true,
    hasBotPlugin = false,
    hasApimPlugin = false,
    envName = "dev"
  ): Promise<[boolean[] | undefined, void, void, boolean, boolean]> {
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

  public async openResourceFolder(): Promise<void> {
    console.log("start to open project: ", this.sampleName);
    const oldPath = path.resolve(this.repoPath, this.originSample);
    // move old sample to project path
    await fs.mkdir(this.projectPath);
    try {
      console.log("oldPath: ", oldPath);
      console.log("newPath: ", this.projectPath);
      await fs.copy(oldPath, this.projectPath);
      await openExistingProject(this.projectPath);
      console.log(
        "create complate !!! ",
        this.appName,
        " path: ",
        this.projectPath
      );
    } catch (error) {
      console.log(error);
      throw new Error(`Failed to open project: ${this.sampleName}`);
    }
  }

  public async createTemplate(): Promise<void> {
    console.log(
      "start to create project: ",
      this.appName,
      ". path: ",
      this.projectPath
    );
    const driver = VSBrowser.instance.driver;
    await clearNotifications();
    await execCommandIfExistFromTreeView(
      TreeViewCommands.CreateProjectCommand,
      Timeout.webView
    );
    const input = await InputBox.create();
    console.log("1. Start from a sample");
    await input.selectQuickPick("Start from a sample");
    await driver.sleep(Timeout.input);
    console.log("2. sample name: ", this.sampleName.split("-").join(" "));
    await input.setText(this.sampleName.split("-").join(" "));
    await input.confirm();
    await driver.sleep(Timeout.input);
    // Input folder path
    await input.setText("default folder");
    const inputbox = await InputBox.create();
    const pick = await inputbox.getQuickPicks();
    const des = (await pick[0].getDescription()) || "";
    this.originPath = path.resolve(des, this.originSample);
    console.log("3. location: ", this.originPath);
    await input.confirm();
    // windows os need to wait long time to create project
    await driver.sleep(Timeout.installWait);

    try {
      console.log(
        "4. copy from origin path",
        this.originPath,
        " to test folder ",
        this.projectPath
      );
      await fs.copy(this.originPath, this.projectPath);
      console.log("copy successfully!!!");
    } catch (error) {
      console.log(error);
      assert.throw(() => error);
    }

    await openExistingProject(this.projectPath);
    console.log(
      "create complate !!! ",
      this.appName,
      " path: ",
      this.projectPath
    );
  }

  public async createTemplateCLI(V3: boolean): Promise<void> {
    console.log(
      "start to create project: ",
      this.appName,
      ". path: ",
      this.projectPath
    );

    await CliHelper.createTemplateProject("./resource", this.originSample, V3);

    this.originPath = path.resolve("./resource", this.originSample);

    try {
      console.log(
        "copy from origin path",
        this.originPath,
        " to test folder ",
        this.projectPath
      );
      await fs.copy(this.originPath, this.projectPath);
      console.log("copy successfully!!!");
    } catch (error) {
      console.log(error);
      assert.throw(() => error);
    }

    await openExistingProject(this.projectPath);
    console.log(
      "create complate !!! ",
      this.appName,
      " path: ",
      this.projectPath
    );
  }

  public async updateManifestAppName(): Promise<void> {
    console.log("[start] update manifest file");
    const manifestFile = fs.pathExistsSync(
      path.resolve(this.projectPath, "appPackage")
    )
      ? path.resolve(this.projectPath, "appPackage", "manifest.json")
      : path.resolve(this.projectPath, "appManifest", "manifest.json");
    try {
      const manifest = await fs.readJSON(manifestFile);
      // manifest name can't be longer than 15 characters
      manifest.name.short =
        this.appName.substring(0, 10) + "${{APP_NAME_SUFFIX}}";
      fs.writeJSON(manifestFile, manifest, { spaces: 4 });
      console.log(
        "[finish] update manifest file successfully, appName: ",
        manifest.name.short
      );
    } catch (error) {
      console.log("[skip] manifest file not found");
    }
  }

  public async openExistFolder(path: string): Promise<void> {
    await openExistingProject(path);
  }

  public async validateLocalStateForBot(): Promise<void> {
    console.log(`validating localSettings for bot`);
    const localStateFile = path.join(".fx", "states", "state.local.json");
    const localStatePath = path.resolve(
      this.testRootFolder,
      this.appName,
      localStateFile
    );
    const localState = await fs.readJSON(localStatePath);

    const botPassword = localState["fx-resource-bot"]["botPassword"] as string;

    const localUserDataFile = path.join(".fx", "states", "local.userdata");
    const localUserDataPath = path.resolve(
      this.testRootFolder,
      this.appName,
      localUserDataFile
    );
    const localUserDataContent = await fs.readFile(localUserDataPath);
    const localUserData = dotenv.parse(localUserDataContent);
    if (
      !localUserData[
        botPassword.substring(2, botPassword.length - 2)
      ].startsWith("crypto_")
    ) {
      throw new Error(`Secret fields are not encrypted for bot project`);
    }
  }

  public async validateLocalStateForTab(): Promise<void> {
    console.log(`validating localSettings for tab`);
    const localStateFile = path.join(".fx", "states", "state.local.json");
    const localStatePath = path.resolve(
      this.testRootFolder,
      this.appName,
      localStateFile
    );
    try {
      await fs.readJSON(localStatePath);
    } catch (error) {
      console.log(error);
      assert.fail(error as string);
    }
  }

  public async provisionWithCLI(
    env: "local" | "dev",
    v3: boolean,
    environment?: NodeJS.ProcessEnv
  ): Promise<void> {
    if (v3) {
      process.env["AZURE_RESOURCE_GROUP_NAME"] = this.rgName;
      await AzSqlHelper.login();
      const azhelper = new AzSqlHelper(this.rgName, []);
      await azhelper.createResourceGroup();
    }
    await CliHelper.provisionProject(this.projectPath, env, v3, environment);
  }

  public async deployWithCLI(env: "local" | "dev"): Promise<void> {
    await CliHelper.deploy(this.projectPath, env);
  }

  public async publishWithCLI(env: "local" | "dev"): Promise<void> {
    await CliHelper.publishProject(this.projectPath, env);
  }

  public async debugWithCLI(env: "local" | "dev", v3?: boolean): Promise<void> {
    await CliHelper.debugProject(this.projectPath, env, v3);
  }

  public async getTeamsAppId(env: "local" | "dev" = "local"): Promise<string> {
    try {
      const userDataFile = path.join(
        TestFilePath.configurationFolder,
        `.env.${env}`
      );
      const configFilePath = path.resolve(this.projectPath, userDataFile);
      const context = dotenvUtil.deserialize(
        await fs.readFile(configFilePath, { encoding: "utf8" })
      );
      const result = context.obj.TEAMS_APP_ID as string;
      console.log(`TEAMS APP ID: ${result}`);
      expect(result).to.not.be.undefined;
      return result;
    } catch (error) {
      console.log(error);
      return "";
    }
  }

  public editDotEnvFile(
    env: "local" | "dev",
    key: string,
    value: string
  ): void {
    const envPath = path.resolve(this.projectPath, "env", `.env.${env}.user`);
    try {
      const envFileContent: string = fs.readFileSync(envPath, "utf-8");
      const envVars: { [key: string]: string } = envFileContent
        .split("\n")
        .reduce((acc: { [key: string]: string }, line: string) => {
          const [key, value] = line.split("=");
          if (key && value) {
            acc[key.trim()] = value.trim();
          }
          return acc;
        }, {});
      envVars[key] = value;
      const newEnvFileContent: string = Object.entries(envVars)
        .map(([key, value]) => `${key}=${value}`)
        .join("\n");
      fs.writeFileSync(envPath, newEnvFileContent);
    } catch (error) {
      console.log('Failed to edit ".env" file.');
    }
  }

  public async prepareDebug(tool: "npm" | "yarn"): Promise<void> {
    {
      console.log(`executor command: npm install yarn`);
      const { stderr, stdout } = await Executor.execute(
        `npm install yarn --force`,
        this.projectPath
      );
      console.log("stdout: ", stdout);
      console.log("stderr: ", stderr);
    }
    {
      console.log(`executor command: corepack enable`);
      const { stderr, stdout } = await Executor.execute(
        `corepack enable`,
        this.projectPath
      );
      console.log("stdout: ", stdout);
      console.log("stderr: ", stderr);
    }
    {
      console.log(`executor command: ${tool} install`);
      const { stderr, stdout } = await Executor.execute(
        `${tool} install`,
        this.projectPath
      );
      console.log("stdout: ", stdout);
      console.log("stderr: ", stderr);
    }
    {
      console.log(`executor command: ${tool} build`);
      const { stderr, stdout } = await Executor.execute(
        `${tool} build`,
        this.projectPath
      );
      console.log("stdout: ", stdout);
      console.log("stderr: ", stderr);
    }
  }

  public async provisionProject(
    appName: string,
    projectPath = "",
    options: {
      createRg?: boolean;
      tool?: "ttk" | "cli";
      option?: string;
      env?: "dev" | "local";
      processEnv?: NodeJS.ProcessEnv;
      skipErrorMessage?: string;
    } = {
      createRg: true,
      tool: "cli",
      option: "",
      env: "dev",
      processEnv: undefined,
      skipErrorMessage: undefined,
    }
  ): Promise<void> {
    if (options.tool === "cli") {
      await this.runCliProvision(projectPath, appName, options);
    } else {
      await runProvision(appName);
    }
  }

  public async deployProject(
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
      await this.runCliDeploy(
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

  public async runCliProvision(
    projectPath: string,
    appName: string,
    options: {
      createRg?: boolean;
      option?: string;
      env?: "dev" | "local";
      processEnv?: NodeJS.ProcessEnv;
      skipErrorMessage?: string;
    } = {
      createRg: true,
      option: "",
      env: "dev",
      processEnv: undefined,
      skipErrorMessage: undefined,
    }
  ): Promise<void> {
    if (options.createRg) {
      console.log("create resource group");
      await createResourceGroup(appName, options.env, "westus");
    }
    const resourceGroupName = `${appName}-${options.env}-rg`;
    process.env["AZURE_RESOURCE_GROUP_NAME"] = resourceGroupName;
    await CliHelper.showVersion(projectPath, options.processEnv);
    const { success, stderr, stdout } = await Executor.provision(
      projectPath,
      options.env,
      true,
      options.skipErrorMessage
    );
    console.log(`stdout: ${stdout}`);
    if (!success) {
      console.log(`stderr: ${stderr}`);
      expect(success).to.be.true;
    }
  }

  public async runCliDeploy(
    projectPath: string,
    option = "",
    env: "dev" | "local" = "dev",
    processEnv?: NodeJS.ProcessEnv,
    retries?: number,
    newCommand?: string
  ) {
    const { success, stderr, stdout } = await Executor.deploy(projectPath, env);
    console.log(`stdout: ${stdout}`);
    if (!success) {
      if (stderr.includes("npm warn")) {
        console.log("[Skip warning]: ", stderr);
      } else {
        console.log(`stderr: ${stderr}`);
        expect(success).to.be.true;
      }
    }
  }

  public createEnvFolder(
    folderPath: string,
    folderName: string
  ): Promise<void> {
    return fs.mkdir(path.resolve(folderPath, folderName));
  }
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as path from "path";
import { openExistingProject } from "../../utils/vscodeOperation";
import * as fs from "fs-extra";
import { execCommand } from "../../utils/execCommand";
import { stopDebugging } from "../../utils/vscodeOperation";
import { TestContext } from "../testContext";
import { dotenvUtil } from "../../utils/envUtil";
import { TestFilePath } from "../../utils/constants";
import { VSBrowser } from "vscode-extension-tester";

export type LocalDebugTestName =
  | "tab"
  | "tabnsso"
  | "function"
  | "bot"
  | "msg"
  | "msgsa"
  | "funcNoti" // functions notification bot
  | "expressnoti" // express notification bot
  | "crbot" // command an response bot
  | "tabbot"
  | "spfx"
  | "spfximport"
  | "botfunc"
  | "template"
  | "m365lp"
  | "workflow"
  | "dashboard"
  | "timeNoti" // timer functions notification bot
  | "ftNoti" // http and timer trigger notification bot
  | "linkunfurl"
  | "aichat"
  | "aiagent"
  | "chatdata"
  | "cdcustomapi"
  | "msgnewapi"
  | "msgapikey"
  | "msgmicroentra";

export class LocalDebugTestContext extends TestContext {
  public testName: LocalDebugTestName;
  public lang: "javascript" | "typescript" | "python";
  public framework: "react" | "minimal" | "none";
  public needMigrate: boolean | undefined;
  public existingSpfxFolder: string;
  public customCopilotRagType: string;
  public customCeopilotAgent: string;
  public llmServiceType: string;

  constructor(
    testName: LocalDebugTestName,
    option?: {
      lang?: "javascript" | "typescript" | "python";
      framework?: "react" | "minimal" | "none";
      needMigrate?: boolean;
      existingSpfxFolder?: string;
      customCopilotRagType?:
        | "custom-copilot-rag-customize"
        | "custom-copilot-rag-azureAISearch"
        | "custom-copilot-rag-customApi"
        | "custom-copilot-rag-microsoft365";
      customCeopilotAgent?:
        | "custom-copilot-agent-new"
        | "custom-copilot-agent-assistants-api";
      llmServiceType?: "llm-service-azure-openai" | "llm-service-openai";
    }
  ) {
    super(testName);
    this.testName = testName;
    this.lang = option?.lang ? option.lang : "javascript";
    this.framework = option?.framework ? option.framework : "react";
    this.needMigrate = option?.needMigrate;
    this.existingSpfxFolder = option?.existingSpfxFolder
      ? option.existingSpfxFolder
      : "existingspfx";
    this.customCopilotRagType = option?.customCopilotRagType
      ? option.customCopilotRagType
      : "custom-copilot-rag-customize";
    this.customCeopilotAgent = option?.customCeopilotAgent
      ? option.customCeopilotAgent
      : "custom-copilot-agent-new";
    this.llmServiceType = option?.llmServiceType
      ? option.llmServiceType
      : "llm-service-azure-openai";
  }

  public async before() {
    await super.before();
    await this.createProject();
    await VSBrowser.instance.driver.sleep(30000);
    // await this.disableDebugConsole();
    const testFolder = path.resolve(this.testRootFolder, this.appName);
    await openExistingProject(testFolder);
  }

  public async after(
    hasAadPlugin = true,
    hasBotPlugin = false,
    hasResourceGroup = false
  ) {
    await stopDebugging();
    await this.context!.close();
    await this.browser!.close();
    await this.cleanResource(
      hasAadPlugin,
      hasBotPlugin,
      "local",
      hasResourceGroup
    );
  }

  public async getTeamsAppId(): Promise<string> {
    const userDataFile = path.join(
      TestFilePath.configurationFolder,
      `.env.local`
    );
    const configFilePath = path.resolve(
      this.testRootFolder,
      this.appName,
      userDataFile
    );
    const context = dotenvUtil.deserialize(
      await fs.readFile(configFilePath, { encoding: "utf8" })
    );
    const result = context.obj.TEAMS_APP_ID as string;
    console.log(`TEAMS APP ID: ${result}`);
    return result;
  }

  public async getM365AppId(): Promise<string> {
    const userDataFile = path.join(
      TestFilePath.configurationFolder,
      `.env.local`
    );
    const configFilePath = path.resolve(
      this.testRootFolder,
      this.appName,
      userDataFile
    );
    const context = dotenvUtil.deserialize(
      await fs.readFile(configFilePath, { encoding: "utf8" })
    );
    const result = context.obj.M365_APP_ID as string;
    console.log(`M365 APP ID: ${result}`);
    return result;
  }

  public async createProject(): Promise<void> {
    if (this.needMigrate) {
      await execCommand(this.testRootFolder, `set TEAMSFX_V3=false`);
    }
    switch (this.testName) {
      case "tab":
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability sso-launch-page --programming-language ${this.lang} --telemetry false`
        );
        break;
      case "tabnsso":
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability tab-non-sso --programming-language ${this.lang} --telemetry false`
        );
        break;
      case "funcNoti":
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability notification --bot-host-type-trigger http-functions --programming-language ${this.lang} --telemetry false`
        );
        break;
      case "expressnoti":
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability notification --bot-host-type-trigger http-express --programming-language ${this.lang} --telemetry false`
        );
        break;
      case "crbot":
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability command-bot --programming-language ${this.lang} --telemetry false`
        );
        break;
      case "function":
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability tab --programming-language ${this.lang} --telemetry false`
        );
        await execCommand(
          path.resolve(this.testRootFolder, this.appName),
          `teamsapp add azure-function --telemetry false`
        );
        break;
      case "bot":
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability bot --programming-language ${this.lang} --telemetry false`
        );
        break;
      case "msg":
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability collect-form-message-extension --programming-language ${this.lang} --telemetry false`
        );
        break;
      case "msgsa":
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability search-app --me-architecture bot-plugin --programming-language ${this.lang} --telemetry false`
        );
        break;
      case "tabbot":
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability tab --programming-language ${this.lang} --telemetry false`
        );
        await execCommand(
          path.resolve(this.testRootFolder, this.appName),
          `teamsapp add bot --telemetry false`
        );
        break;
      case "spfx":
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability tab-spfx --spfx-framework-type ${this.framework} --spfx-webpart-name ${this.appName} --telemetry false`
        );
        break;
      case "spfximport":
        const resourcePath = path.resolve(
          __dirname,
          "../../../.test-resources/",
          this.existingSpfxFolder
        );
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability tab-spfx --spfx-solution import --spfx-folder ${resourcePath} --telemetry false`
        );
        break;
      case "botfunc":
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability tab --programming-language ${this.lang} --telemetry false`
        );
        await execCommand(
          path.resolve(this.testRootFolder, this.appName),
          `teamsapp add azure-function --telemetry false`
        );
        await execCommand(
          path.resolve(this.testRootFolder, this.appName),
          `teamsapp add bot --telemetry false`
        );
        break;
      case "m365lp":
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability sso-launch-page --programming-language ${this.lang} --telemetry false`
        );
        break;
      case "workflow":
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability workflow-bot --programming-language ${this.lang} --telemetry false`
        );
        break;
      case "dashboard":
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability dashboard-tab --programming-language ${this.lang} --telemetry false`
        );
        break;
      case "timeNoti":
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability notification --bot-host-type-trigger timer-functions --programming-language ${this.lang} --telemetry false`
        );
        break;
      case "ftNoti":
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability notification --bot-host-type-trigger http-and-timer-functions --programming-language ${this.lang} --telemetry false`
        );
        break;
      case "linkunfurl":
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability link-unfurling  --programming-language ${this.lang} --telemetry false`
        );
        break;
      case "aichat":
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability custom-copilot-basic --llm-service ${this.llmServiceType} --programming-language ${this.lang} --telemetry false`
        );
        break;
      case "aiagent":
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability custom-copilot-agent --custom-copilot-agent ${this.customCeopilotAgent} --llm-service ${this.llmServiceType} --programming-language ${this.lang} --telemetry false`
        );
        break;
      case "chatdata":
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability custom-copilot-rag --custom-copilot-rag ${this.customCopilotRagType} --llm-service ${this.llmServiceType} --programming-language ${this.lang} --telemetry false`
        );
        break;
      case "msgnewapi":
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability search-app  --me-architecture new-api --programming-language ${this.lang} --telemetry false`
        );
        break;
      case "msgapikey":
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability search-app  --me-architecture new-api --api-auth bearer-token --programming-language ${this.lang} --telemetry false`
        );
        break;
      case "msgmicroentra":
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability search-app  --me-architecture new-api --api-auth microsoft-entra --programming-language ${this.lang} --telemetry false`
        );
        break;
      case "cdcustomapi": //chat data customApi
        const apiSpecPath =
          "https://raw.githubusercontent.com/SLdragon/example-openapi-spec/main/real-no-auth.yaml";
        await execCommand(
          this.testRootFolder,
          `teamsapp new --app-name ${this.appName} --interactive false --capability custom-copilot-rag --custom-copilot-rag ${this.customCopilotRagType} --llm-service ${this.llmServiceType} --programming-language ${this.lang}  --openapi-spec-location ${apiSpecPath} --api-operation "GET /repairs" --telemetry false`
        );
        break;
    }
    if (this.needMigrate) {
      await execCommand(this.testRootFolder, `set TEAMSFX_V3=true`);
    }
  }

  public async disableDebugConsole(): Promise<void> {
    const filePath = path.resolve(
      this.testRootFolder,
      this.appName,
      ".vscode/launch.json"
    );
    const content = await fs.readJson(filePath);
    const configs = content.configurations as any[];
    for (const config of configs) {
      config.internalConsoleOptions = "neverOpen";
    }
    await fs.writeJson(filePath, content);
  }

  // related bug https://msazure.visualstudio.com/Microsoft%20Teams%20Extensibility/_workitems/edit/12850424
  public async validateLocalStateForTab(): Promise<any> {
    console.log(`validating localSettings for tab`);
    const localStateFile = path.join(
      TestFilePath.configurationFolder,
      ".env.local.user"
    );
    const localStatePath = path.resolve(
      this.testRootFolder,
      this.appName,
      localStateFile
    );

    const localState = dotenvUtil.deserialize(
      await fs.readFile(localStatePath, { encoding: "utf8" })
    );

    const aadPassword = localState.obj.SECRET_AAD_APP_CLIENT_SECRET as string;

    if (!aadPassword.startsWith("crypto_")) {
      throw new Error(`Secret fields are not encrypted for sso aad project`);
    }
  }

  // releated bug https://msazure.visualstudio.com/Microsoft%20Teams%20Extensibility/_workitems/edit/12856828
  public async validateLocalStateForBot(): Promise<any> {
    console.log(`validating localSettings for bot`);
    const localStateFile = path.join(
      TestFilePath.configurationFolder,
      ".env.local.user"
    );
    const localStatePath = path.resolve(
      this.testRootFolder,
      this.appName,
      localStateFile
    );

    const localState = dotenvUtil.deserialize(
      await fs.readFile(localStatePath, { encoding: "utf8" })
    );

    const botPassword = localState.obj.SECRET_BOT_PASSWORD as string;

    if (!botPassword.startsWith("crypto_")) {
      throw new Error(`Secret fields are not encrypted for bot project`);
    }
  }
}

export class LocalDebugSampleTestContext extends LocalDebugTestContext {
  public sampleName?: string;
  constructor(sampleName: string) {
    super("template");
    this.testName = "template";
    this.sampleName = sampleName;
  }
}

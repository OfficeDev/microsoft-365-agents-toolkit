// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */
import * as path from "path";
import { VSBrowser } from "vscode-extension-tester";
import {
  startDebugging,
  waitForTerminal,
  execCommandIfExist,
  createNewProject,
} from "../../../utils/vscodeOperation";
import {
  initPage,
  validateWelcomeAndReplyBot,
} from "../../../utils/playwrightOperation";
import { LocalDebugTestContext } from "../../localdebug/localdebugContext";
import {
  RemoteDebugTestContext,
  deployProject,
  provisionProject,
} from "../../remotedebug/remotedebugContext";
import {
  Timeout,
  LocalDebugTaskLabel,
  DebugItemSelect,
  ValidationContent,
  Lang,
  LocalDebugTaskResult,
} from "../../../utils/constants";
import { Env, OpenAiKey } from "../../../utils/env";
import { it } from "../../../utils/it";
import { editDotEnvFile, validateFileExist } from "../../../utils/commonUtils";

export function happyPathTest(options: {
  lang: Lang;
  llm: "llm-service-openai" | "llm-service-azure-openai";
  testPlanCaseId_local?: number;
  testPlanCaseId_dev?: number;
  author: string;
}): void {
  describe("Debug Tests", function () {
    this.timeout(Timeout.testCase);
    let localDebugTestContext: LocalDebugTestContext;
    let remoteDebugTestContext: RemoteDebugTestContext;
    let testRootFolder: string;
    let appName: string;
    const appNameCopySuffix = "copy";
    let newAppFolderName: string;
    let projectPath: string;
    let debugContent: "local" | "remote" | undefined = undefined;

    beforeEach(async function () {
      if (debugContent === undefined) {
        debugContent = "remote";
      } else {
        debugContent = "local";
      }

      // ensure workbench is ready
      this.timeout(Timeout.prepareTestCase);
      if (debugContent === "local") {
        localDebugTestContext = new LocalDebugTestContext("weather", {
          lang: options.lang,
          llmServiceType: options.llm,
        });
        await localDebugTestContext.before();
      } else {
        remoteDebugTestContext = new RemoteDebugTestContext("weather");
        testRootFolder = remoteDebugTestContext.testRootFolder;
        appName = remoteDebugTestContext.appName;
        newAppFolderName = appName + appNameCopySuffix;
        projectPath = path.resolve(testRootFolder, newAppFolderName);
        await remoteDebugTestContext.before();
      }
    });

    afterEach(async function () {
      this.timeout(Timeout.finishTestCase);
      if (debugContent === "local") {
        await localDebugTestContext.after(false, true);
      } else {
        await remoteDebugTestContext.after();

        //Close the folder and cleanup local sample project
        await execCommandIfExist(
          "Workspaces: Close Workspace",
          Timeout.webView
        );
        console.log(`[Successfully] start to clean up for ${projectPath}`);
        await remoteDebugTestContext.cleanUp(
          appName,
          projectPath,
          false,
          true,
          false
        );
      }
    });

    it(
      `[auto][${options.lang}][${
        options.llm === "llm-service-azure-openai" ? "Azure OpenAI" : "OpenAI"
      }] Remote debug for Weather Agent - ${
        options.lang === "JavaScript" ? "JavaScript" : "TypeScript"
      }`,
      {
        testPlanCaseId: options.testPlanCaseId_dev,
        author: options.author,
      },
      async function () {
        const driver = VSBrowser.instance.driver;
        await createNewProject("weather", appName, {
          lang: options.lang,
          aiType:
            options.llm === "llm-service-azure-openai"
              ? "Azure OpenAI"
              : "OpenAI",
        });
        validateFileExist(
          projectPath,
          `src/${options.lang === Lang.JS ? "index.js" : "index.ts"}`
        );

        const envPath = path.resolve(projectPath, "env", ".env.dev.user");
        let isRealKey = false;
        if (options.llm === "llm-service-azure-openai") {
          isRealKey = OpenAiKey.azureOpenAiKey ? true : false;
          const azureOpenAiKey = OpenAiKey.azureOpenAiKey
            ? OpenAiKey.azureOpenAiKey
            : "fake";
          const azureOpenAiEndpoint = OpenAiKey.azureOpenAiEndpoint
            ? OpenAiKey.azureOpenAiEndpoint
            : "https://test.com";
          const azureOpenAiModelDeploymentName =
            OpenAiKey.azureOpenAiModelDeploymentName
              ? OpenAiKey.azureOpenAiModelDeploymentName
              : "fake";
          editDotEnvFile(
            envPath,
            "SECRET_AZURE_OPENAI_API_KEY",
            azureOpenAiKey
          );
          editDotEnvFile(envPath, "AZURE_OPENAI_ENDPOINT", azureOpenAiEndpoint);
          editDotEnvFile(
            envPath,
            "AZURE_OPENAI_DEPLOYMENT_NAME",
            azureOpenAiModelDeploymentName
          );
        } else {
          // openai entrance
          editDotEnvFile(envPath, "SECRET_OPENAI_API_KEY", "fake");
        }

        {
          // create azure assistant need to use local env
          const localEnvPath = path.resolve(
            projectPath,
            "env",
            ".env.local.user"
          );
          const azureOpenAiKey = OpenAiKey.azureOpenAiKey
            ? OpenAiKey.azureOpenAiKey
            : "fake";
          const azureOpenAiEndpoint = OpenAiKey.azureOpenAiEndpoint
            ? OpenAiKey.azureOpenAiEndpoint
            : "https://test.com";
          const azureOpenAiModelDeploymentName =
            OpenAiKey.azureOpenAiModelDeploymentName
              ? OpenAiKey.azureOpenAiModelDeploymentName
              : "fake";
          editDotEnvFile(
            localEnvPath,
            "SECRET_AZURE_OPENAI_API_KEY",
            azureOpenAiKey
          );
          editDotEnvFile(
            localEnvPath,
            "AZURE_OPENAI_ENDPOINT",
            azureOpenAiEndpoint
          );
          editDotEnvFile(
            localEnvPath,
            "AZURE_OPENAI_DEPLOYMENT_NAME",
            azureOpenAiModelDeploymentName
          );
        }

        await provisionProject(appName, projectPath);
        await deployProject(projectPath, Timeout.botDeploy);
        await driver.sleep(Timeout.shortTimeLoading);
        // [known issue] python remote need deploy twice
        await deployProject(projectPath, Timeout.botDeploy);
        await driver.sleep(Timeout.tabDeploy);
        const teamsAppId = await remoteDebugTestContext.getTeamsAppId(
          projectPath
        );

        const page = await initPage(
          remoteDebugTestContext.context!,
          teamsAppId,
          Env.username,
          Env.password,
          {
            projectPath: projectPath,
            env: "dev",
            teamsAppName: appName,
            searchApp: false,
          }
        );

        if (isRealKey) {
          await validateWelcomeAndReplyBot(page, {
            hasWelcomeMessage: true,
            hasCommandReplyValidation: true,
            botCommand:
              "Can you forecast the tomorrow weather in San Francisco for me?",
            expectedWelcomeMessage:
              ValidationContent.WeatherBotMessageWelcomeInstruction,
            expectedReplyMessage: ValidationContent.WeatherBotMessage,
            timeout: Timeout.longTimeWait,
          });
        } else {
          await validateWelcomeAndReplyBot(page, {
            hasWelcomeMessage: false,
            hasCommandReplyValidation: true,
            botCommand:
              "Can you forecast the tomorrow weather in San Francisco for me?",
            expectedWelcomeMessage:
              ValidationContent.WeatherBotMessageWelcomeInstruction,
            expectedReplyMessage: ValidationContent.AiBotErrorMessage,
            timeout: Timeout.longTimeWait,
          });
        }
      }
    );

    it(
      `[auto][${options.lang}][${
        options.llm === "llm-service-azure-openai" ? "Azure OpenAI" : "OpenAI"
      }] Local debug for AI Agent - ${
        options.lang === "JavaScript" ? "JavaScript" : "TypeScript"
      }`,
      {
        testPlanCaseId: options.testPlanCaseId_local,
        author: options.author,
      },
      async function () {
        const projectPath = path.resolve(
          localDebugTestContext.testRootFolder,
          localDebugTestContext.appName
        );
        validateFileExist(
          projectPath,
          `src/${options.lang === Lang.JS ? "index.js" : "index.ts"}`
        );

        const envPath = path.resolve(projectPath, "env", ".env.local.user");
        // azure openai entrance
        let isRealKey = false;
        if (options.llm === "llm-service-azure-openai") {
          isRealKey = OpenAiKey.azureOpenAiKey ? true : false;
          const azureOpenAiKey = OpenAiKey.azureOpenAiKey
            ? OpenAiKey.azureOpenAiKey
            : "fake";
          const azureOpenAiEndpoint = OpenAiKey.azureOpenAiEndpoint
            ? OpenAiKey.azureOpenAiEndpoint
            : "https://test.com";
          const azureOpenAiModelDeploymentName =
            OpenAiKey.azureOpenAiModelDeploymentName
              ? OpenAiKey.azureOpenAiModelDeploymentName
              : "fake";
          editDotEnvFile(
            envPath,
            "SECRET_AZURE_OPENAI_API_KEY",
            azureOpenAiKey
          );
          editDotEnvFile(envPath, "AZURE_OPENAI_ENDPOINT", azureOpenAiEndpoint);
          editDotEnvFile(
            envPath,
            "AZURE_OPENAI_DEPLOYMENT_NAME",
            azureOpenAiModelDeploymentName
          );
        } else {
          // openai entrance
          editDotEnvFile(envPath, "SECRET_OPENAI_API_KEY", "fake");
        }

        await startDebugging(DebugItemSelect.DebugInTeamsUsingChrome);

        await waitForTerminal(LocalDebugTaskLabel.StartLocalTunnel);
        await waitForTerminal(
          LocalDebugTaskLabel.StartBotApp,
          LocalDebugTaskResult.DebuggerAttached
        );

        const teamsAppId = await localDebugTestContext.getTeamsAppId();
        const page = await initPage(
          localDebugTestContext.context!,
          teamsAppId,
          Env.username,
          Env.password,
          {
            projectPath: projectPath,
            teamsAppName: localDebugTestContext.appName,
            env: "local",
            searchApp: false,
          }
        );
        if (isRealKey) {
          await validateWelcomeAndReplyBot(page, {
            hasWelcomeMessage: true,
            hasCommandReplyValidation: true,
            botCommand:
              "Can you forecast the tomorrow weather in San Francisco for me?",
            expectedWelcomeMessage:
              ValidationContent.WeatherBotMessageWelcomeInstruction,
            expectedReplyMessage: ValidationContent.WeatherBotMessage,
            timeout: Timeout.longTimeWait,
          });
        } else {
          await validateWelcomeAndReplyBot(page, {
            hasWelcomeMessage: false,
            hasCommandReplyValidation: true,
            botCommand:
              "Can you forecast the tomorrow weather in San Francisco for me?",
            expectedWelcomeMessage:
              ValidationContent.WeatherBotMessageWelcomeInstruction,
            expectedReplyMessage: ValidationContent.AiBotErrorMessage,
            timeout: Timeout.longTimeWait,
          });
        }
      }
    );
  });
}

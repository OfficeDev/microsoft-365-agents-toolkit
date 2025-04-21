// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Anne Fu <v-annefu@microsoft.com>
 */
import * as path from "path";
import { VSBrowser } from "vscode-extension-tester";
import { Lang, Timeout, ValidationContent } from "../../utils/constants";
import {
  RemoteDebugTestContext,
  provisionProject,
  deployProject,
} from "./remotedebugContext";
import {
  execCommandIfExist,
  createNewProject,
} from "../../utils/vscodeOperation";
import { initPage, validateCustomapi } from "../../utils/playwrightOperation";
import { Env, OpenAiKey } from "../../utils/env";
import { it } from "../../utils/it";
import { editDotEnvFile, validateFileExist } from "../../utils/commonUtils";
import { RetryHandler } from "../../utils/retryHandler";

describe("Remote debug Tests", function () {
  this.timeout(Timeout.testAzureCase);
  let remoteDebugTestContext: RemoteDebugTestContext;
  let testRootFolder: string;
  let appName: string;
  const appNameCopySuffix = "copy";
  let newAppFolderName: string;
  let projectPath: string;

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    remoteDebugTestContext = new RemoteDebugTestContext("cdcustomapi");
    testRootFolder = remoteDebugTestContext.testRootFolder;
    appName = remoteDebugTestContext.appName;
    newAppFolderName = appName + appNameCopySuffix;
    projectPath = path.resolve(testRootFolder, newAppFolderName);
    await remoteDebugTestContext.before();
  });

  afterEach(async function () {
    this.timeout(Timeout.finishAzureTestCase);
    await remoteDebugTestContext.after();

    //Close the folder and cleanup local sample project
    await execCommandIfExist("Workspaces: Close Workspace", Timeout.webView);
    console.log(`[Successfully] start to clean up for ${projectPath}`);
    await remoteDebugTestContext.cleanUp(
      appName,
      projectPath,
      false,
      true,
      false
    );
  });

  it(
    "[auto][Python][OpenAI] Remote debug for Custom Copilot Rag Custom Api",
    {
      testPlanCaseId: 29165762,
      author: "v-annefu@microsoft.com",
    },
    async function () {
      const driver = VSBrowser.instance.driver;
      await createNewProject("cdcustomapi", appName, {
        aiType: "OpenAI",
        lang: Lang.PY,
        dataOption: "Custom API",
      });
      validateFileExist(projectPath, "src/app.py");
      const envPath = path.resolve(projectPath, "env", ".env.dev.user");
      const isRealKey = OpenAiKey.openAiKey ? true : false;
      const openAiKey = OpenAiKey.openAiKey ? OpenAiKey.openAiKey : "fake";
      editDotEnvFile(envPath, "SECRET_OPENAI_API_KEY", openAiKey);
      await provisionProject(appName, projectPath);
      await deployProject(projectPath, Timeout.botDeploy);
      const teamsAppId = await remoteDebugTestContext.getTeamsAppId(
        projectPath
      );
      const page = await initPage(
        remoteDebugTestContext.context!,
        teamsAppId,
        Env.username,
        Env.password,
        { projectPath: projectPath, env: "dev" }
      );
      await driver.sleep(Timeout.longTimeWait);
      try {
        if (isRealKey) {
          await validateCustomapi(page, {
            hasWelcomeMessage: false,
            hasCommandReplyValidation: true,
            botCommand: "Get repairs for Karin",
            expectedWelcomeMessage:
              ValidationContent.AiChatBotWelcomeInstruction,
            expectedReplyMessage: "assignedTo: Karin",
            timeout: Timeout.longTimeWait,
          });
        } else {
          await validateCustomapi(page, {
            hasWelcomeMessage: false,
            hasCommandReplyValidation: true,
            botCommand: "helloWorld",
            expectedWelcomeMessage:
              ValidationContent.AiChatBotWelcomeInstruction,
            expectedReplyMessage: ValidationContent.AiBotErrorMessage,
            timeout: Timeout.longTimeWait,
          });
        }
      } catch {
        await RetryHandler.retry(async () => {
          await deployProject(projectPath, Timeout.botDeploy);
          await driver.sleep(Timeout.longTimeWait);
          if (isRealKey) {
            await validateCustomapi(page, {
              hasWelcomeMessage: false,
              hasCommandReplyValidation: true,
              botCommand: "Get repairs for Karin",
              expectedWelcomeMessage:
                ValidationContent.AiChatBotWelcomeInstruction,
              expectedReplyMessage: "assignedTo: Karin",
              timeout: Timeout.longTimeWait,
            });
          } else {
            await validateCustomapi(page, {
              hasWelcomeMessage: false,
              hasCommandReplyValidation: true,
              botCommand: "helloWorld",
              expectedWelcomeMessage:
                ValidationContent.AiChatBotWelcomeInstruction,
              expectedReplyMessage: ValidationContent.AiBotErrorMessage,
              timeout: Timeout.longTimeWait,
            });
          }
        }, 2);
      }
    }
  );
});

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */
import * as path from "path";
import { VSBrowser } from "vscode-extension-tester";
import { Timeout, ValidationContent } from "../../utils/constants";
import {
  RemoteDebugTestContext,
  provisionProject,
  deployProject,
} from "./remotedebugContext";
import {
  execCommandIfExist,
  createNewProject,
} from "../../utils/vscodeOperation";
import {
  initPage,
  validateWelcomeAndReplyBot,
} from "../../utils/playwrightOperation";
import { Env, OpenAiKey } from "../../utils/env";
import { it } from "../../utils/it";
import { editDotEnvFile, validateFileExist } from "../../utils/commonUtils";
import { AzSearchHelper } from "../../utils/azureCliHelper";
import { Executor } from "../../utils/executor";

describe("Remote debug Tests", function () {
  this.timeout(Timeout.testAzureCase);
  let remoteDebugTestContext: RemoteDebugTestContext;
  let testRootFolder: string;
  let appName: string;
  const appNameCopySuffix = "copy";
  let newAppFolderName: string;
  let projectPath: string;
  let azSearchHelper: AzSearchHelper;

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    remoteDebugTestContext = new RemoteDebugTestContext("chatdata");
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
    "[auto][TS][OpenAI] Remote debug for basic rag bot using azure ai search data",
    {
      testPlanCaseId: 28970327,
      author: "v-ivanchen@microsoft.com",
    },
    async function () {
      const driver = VSBrowser.instance.driver;
      await createNewProject("chatdata", appName, {
        lang: "TypeScript",
        aiType: "OpenAI",
        dataOption: "Azure AI Search",
      });
      validateFileExist(projectPath, "src/index.ts");
      const envPath = path.resolve(projectPath, "env", ".env.dev.user");

      const searchKey = "fake";
      const searchEndpoint = "https://test.com";
      const openAiKey = "fake";
      editDotEnvFile(envPath, "SECRET_OPENAI_API_KEY", openAiKey);
      editDotEnvFile(envPath, "SECRET_AZURE_SEARCH_KEY", searchKey);
      editDotEnvFile(envPath, "AZURE_SEARCH_ENDPOINT", searchEndpoint);

      // prepare for the npm run indexer:create
      const testToolEnvPath = path.resolve(
        projectPath,
        "env",
        ".env.testtool.user"
      );
      editDotEnvFile(testToolEnvPath, "SECRET_OPENAI_API_KEY", openAiKey);
      editDotEnvFile(testToolEnvPath, "SECRET_AZURE_SEARCH_KEY", searchKey);
      editDotEnvFile(testToolEnvPath, "AZURE_SEARCH_ENDPOINT", searchEndpoint);

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

      await validateWelcomeAndReplyBot(page, {
        hasWelcomeMessage: false,
        hasCommandReplyValidation: true,
        botCommand: "helloWorld",
        expectedWelcomeMessage: ValidationContent.AiChatBotWelcomeInstruction,
        expectedReplyMessage: ValidationContent.AiBotErrorMessage,
        timeout: Timeout.longTimeWait,
      });
    }
  );
});

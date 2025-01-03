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
import { initPage, validateBot } from "../../utils/playwrightOperation";
import { Env, OpenAiKey } from "../../utils/env";
import { it } from "../../utils/it";
import { editDotEnvFile, validateFileExist } from "../../utils/commonUtils";

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
    "[auto][TS][OpenAI] Remote debug for basic rag bot using m365 data",
    {
      testPlanCaseId: 29022981,
      author: "v-ivanchen@microsoft.com",
    },
    async function () {
      const driver = VSBrowser.instance.driver;
      await createNewProject("chatdata", appName, {
        lang: "TypeScript",
        aiType: "OpenAI",
        dataOption: "Microsoft 365",
      });
      validateFileExist(projectPath, "src/index.ts");
      const envPath = path.resolve(projectPath, "env", ".env.dev.user");
      const isRealKey = false;
      const openAiKey = "fake";
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
        Env.password
      );
      await driver.sleep(Timeout.longTimeWait);
      console.log("consent login");
      await validateBot(page, {
        botCommand: "show",
        expected: "You are successfully logged in.",
        consentPrompt: true,
      });
      console.log("validate bot message");
      if (isRealKey) {
        await validateBot(page, {
          botCommand: "Tell me about Contoso Electronics history",
          expected: "1985",
          consentPrompt: false,
        });
      } else {
        try {
          await validateBot(page, {
            botCommand: "Tell me about Contoso Electronics history",
            expected: "1985",
            consentPrompt: false,
          });
        } catch (error) {
          await validateBot(page, {
            botCommand: "Tell me about Contoso Electronics history",
            expected: ValidationContent.AiBotErrorMessage,
            consentPrompt: false,
          });
        }
      }
    }
  );
});

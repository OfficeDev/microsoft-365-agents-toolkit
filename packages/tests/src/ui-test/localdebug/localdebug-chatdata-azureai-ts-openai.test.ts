// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */
import * as path from "path";
import { startDebugging, waitForTerminal } from "../../utils/vscodeOperation";
import {
  initPage,
  validateWelcomeAndReplyBot,
} from "../../utils/playwrightOperation";
import { LocalDebugTestContext } from "./localdebugContext";
import {
  Timeout,
  LocalDebugTaskLabel,
  DebugItemSelect,
  ValidationContent,
} from "../../utils/constants";
import { Env } from "../../utils/env";
import { it } from "../../utils/it";
import { editDotEnvFile, validateFileExist } from "../../utils/commonUtils";
import { AzSearchHelper } from "../../utils/azureCliHelper";

describe("Local Debug Tests", function () {
  this.timeout(Timeout.testCase);
  let localDebugTestContext: LocalDebugTestContext;
  let azSearchHelper: AzSearchHelper;

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    localDebugTestContext = new LocalDebugTestContext("chatdata", {
      lang: "typescript",
      customCopilotRagType: "custom-copilot-rag-azureAISearch",
      llmServiceType: "llm-service-openai",
    });
    await localDebugTestContext.before();
  });

  afterEach(async function () {
    this.timeout(Timeout.finishTestCase);
    await localDebugTestContext.after(false, true, true);
  });

  it(
    "[auto][TS][OpenAI] Local debug for basic rag bot using azure ai data",
    {
      testPlanCaseId: 28970306,
      author: "v-ivanchen@microsoft.com",
    },
    async function () {
      const projectPath = path.resolve(
        localDebugTestContext.testRootFolder,
        localDebugTestContext.appName
      );
      validateFileExist(projectPath, "src/index.ts");
      const envPath = path.resolve(projectPath, "env", ".env.local.user");

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

      await startDebugging(DebugItemSelect.DebugInTeamsUsingChrome);
      await waitForTerminal(LocalDebugTaskLabel.StartLocalTunnel);
      await waitForTerminal(LocalDebugTaskLabel.StartBotApp, "Bot Started");

      const teamsAppId = await localDebugTestContext.getTeamsAppId();
      const page = await initPage(
        localDebugTestContext.context!,
        teamsAppId,
        Env.username,
        Env.password
      );
      await localDebugTestContext.validateLocalStateForBot();

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

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */
import * as path from "path";
import {
  createEnvironmentWithPython,
  startDebugging,
  waitForTerminal,
} from "../../utils/vscodeOperation";
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
  LocalDebugTaskLabel2,
  Lang,
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
      lang: Lang.PY,
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
    "[auto][Python][OpenAI] Local debug for basic rag bot using azure ai data",
    {
      testPlanCaseId: 27454158,
      author: "v-ivanchen@microsoft.com",
    },
    async function () {
      const projectPath = path.resolve(
        localDebugTestContext.testRootFolder,
        localDebugTestContext.appName
      );
      validateFileExist(projectPath, "src/app.py");
      const envPath = path.resolve(projectPath, "env", ".env.local.user");

      const searchKey = "fake";
      const searchEndpoint = "https://test.com";
      const openAiKey = "fake";
      editDotEnvFile(envPath, "SECRET_OPENAI_API_KEY", openAiKey);
      editDotEnvFile(envPath, "SECRET_AZURE_SEARCH_KEY", searchKey);
      editDotEnvFile(envPath, "AZURE_SEARCH_ENDPOINT", searchEndpoint);

      await createEnvironmentWithPython();
      // create azure search data

      await startDebugging(DebugItemSelect.DebugInTeamsUsingChrome);

      await waitForTerminal(LocalDebugTaskLabel.StartLocalTunnel);
      await waitForTerminal(
        LocalDebugTaskLabel2.PythonDebugConsole,
        "Running on http://localhost:3978"
      );

      const teamsAppId = await localDebugTestContext.getTeamsAppId();
      const page = await initPage(
        localDebugTestContext.context!,
        teamsAppId,
        Env.username,
        Env.password,
        { projectPath: projectPath, env: "local" }
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

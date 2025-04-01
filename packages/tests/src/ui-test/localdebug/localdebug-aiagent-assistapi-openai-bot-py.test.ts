// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Helly Zhang <v-helzha@microsoft.com>
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
} from "../../utils/constants";
import { Env, OpenAiKey } from "../../utils/env";
import { it } from "../../utils/it";
import { editDotEnvFile, validateFileExist } from "../../utils/commonUtils";

describe("Local Debug Tests", function () {
  this.timeout(Timeout.testCase);
  let localDebugTestContext: LocalDebugTestContext;

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    localDebugTestContext = new LocalDebugTestContext("aiagent", {
      lang: "python",
      customCeopilotAgent: "custom-copilot-agent-assistants-api",
      llmServiceType: "llm-service-openai",
    });
    await localDebugTestContext.before();
  });

  afterEach(async function () {
    this.timeout(Timeout.finishTestCase);
    await localDebugTestContext.after(false, true);
  });

  it(
    "[auto][Python][OpenAI] Local debug for AI Agent - Build with Assistants API",
    {
      testPlanCaseId: 28165244,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      const projectPath = path.resolve(
        localDebugTestContext.testRootFolder,
        localDebugTestContext.appName
      );
      validateFileExist(projectPath, "src/app.py");
      const envPath = path.resolve(projectPath, "env", ".env.local.user");
      editDotEnvFile(envPath, "SECRET_OPENAI_API_KEY", "fake");
      editDotEnvFile(envPath, "OPENAI_ASSISTANT_ID", "fake");

      await createEnvironmentWithPython();

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
        expectedWelcomeMessage:
          ValidationContent.AiAssistantBotWelcomeInstruction,
        expectedReplyMessage: ValidationContent.AiBotErrorMessage,
        timeout: Timeout.longTimeWait,
      });
    }
  );
});

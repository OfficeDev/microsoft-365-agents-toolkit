// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */
import * as path from "path";
import { startDebugging, waitForTerminal } from "../../utils/vscodeOperation";
import { initPage, validateBot } from "../../utils/playwrightOperation";
import { LocalDebugTestContext } from "./localdebugContext";
import {
  Timeout,
  LocalDebugTaskLabel,
  DebugItemSelect,
  ValidationContent,
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
    localDebugTestContext = new LocalDebugTestContext("chatdata", {
      customCopilotRagType: "custom-copilot-rag-microsoft365",
    });
    await localDebugTestContext.before();
  });

  afterEach(async function () {
    this.timeout(Timeout.finishTestCase);
    await localDebugTestContext.after(false, true, true);
  });

  it(
    "[auto][JS][Azure OpenAI] Local debug for basic rag bot using m365 ai data",
    {
      testPlanCaseId: 27569160,
      author: "v-ivanchen@microsoft.com",
    },
    async function () {
      const projectPath = path.resolve(
        localDebugTestContext.testRootFolder,
        localDebugTestContext.appName
      );
      validateFileExist(projectPath, "src/index.js");
      const envPath = path.resolve(projectPath, "env", ".env.local.user");

      const isRealKey = OpenAiKey.azureOpenAiKey ? true : false;
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

      editDotEnvFile(envPath, "SECRET_AZURE_OPENAI_API_KEY", azureOpenAiKey);
      editDotEnvFile(envPath, "AZURE_OPENAI_ENDPOINT", azureOpenAiEndpoint);
      editDotEnvFile(
        envPath,
        "AZURE_OPENAI_DEPLOYMENT_NAME",
        azureOpenAiModelDeploymentName
      );

      console.log(`
        SECRET_AZURE_OPENAI_API_KEY=${azureOpenAiKey}
        AZURE_OPENAI_ENDPOINT=${azureOpenAiEndpoint}
        AZURE_OPENAI_DEPLOYMENT_NAME=${azureOpenAiModelDeploymentName}
      `);

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
        await validateBot(page, {
          botCommand: "Tell me about Contoso Electronics history",
          expected: ValidationContent.AiBotErrorMessage,
          consentPrompt: false,
        });
      }
    }
  );
});

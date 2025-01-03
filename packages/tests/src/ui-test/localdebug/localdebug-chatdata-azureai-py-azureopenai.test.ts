// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */
import * as path from "path";
import * as fs from "fs";
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
import { AzSearchHelper } from "../../utils/azureCliHelper";
import { Executor } from "../../utils/executor";

describe("Local Debug Tests", function () {
  this.timeout(Timeout.testCase);
  let localDebugTestContext: LocalDebugTestContext;
  let azSearchHelper: AzSearchHelper;

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    localDebugTestContext = new LocalDebugTestContext("chatdata", {
      lang: "python",
      customCopilotRagType: "custom-copilot-rag-azureAISearch",
    });
    await localDebugTestContext.before();
  });

  afterEach(async function () {
    this.timeout(Timeout.finishTestCase);
    await localDebugTestContext.after(false, true, true);
  });

  it(
    "[auto][Python][Azure OpenAI] Local debug for basic rag bot using azure ai data",
    {
      testPlanCaseId: 27454153,
      author: "v-ivanchen@microsoft.com",
    },
    async function () {
      const projectPath = path.resolve(
        localDebugTestContext.testRootFolder,
        localDebugTestContext.appName
      );
      validateFileExist(projectPath, "src/app.py");
      const envPath = path.resolve(projectPath, "env", ".env.local.user");

      const isRealKey = OpenAiKey.azureOpenAiKey ? true : false;
      // create azure search
      if (isRealKey) {
        const rgName = `${localDebugTestContext.appName}-local-rg`;

        azSearchHelper = new AzSearchHelper(rgName);
        await azSearchHelper.createSearch();
      }
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
        "AZURE_OPENAI_MODEL_DEPLOYMENT_NAME",
        azureOpenAiModelDeploymentName
      );
      const embeddingDeploymentName =
        OpenAiKey.azureOpenAiEmbeddingDeploymentName ?? "fake";
      editDotEnvFile(
        envPath,
        "AZURE_OPENAI_EMBEDDING_DEPLOYMENT",
        embeddingDeploymentName
      );
      const searchKey = isRealKey ? azSearchHelper.apiKey : "fake";
      const searchEndpoint = isRealKey
        ? azSearchHelper.endpoint
        : "https://test.com";
      editDotEnvFile(envPath, "SECRET_AZURE_SEARCH_KEY", searchKey);
      editDotEnvFile(envPath, "AZURE_SEARCH_ENDPOINT", searchEndpoint);

      console.log(`
        SECRET_AZURE_OPENAI_API_KEY=${azureOpenAiKey}
        AZURE_OPENAI_ENDPOINT=${azureOpenAiEndpoint}
        AZURE_OPENAI_DEPLOYMENT_NAME=${azureOpenAiModelDeploymentName}
        AZURE_OPENAI_EMBEDDING_DEPLOYMENT=${embeddingDeploymentName}
        SECRET_AZURE_SEARCH_KEY=${searchKey}
        AZURE_SEARCH_ENDPOINT=${searchEndpoint}
      `);

      await createEnvironmentWithPython();
      // create azure search data
      if (isRealKey) {
        console.log("Start to create azure search data");
        const installCmd = `python src/indexers/setup.py --api-key ${azureOpenAiKey} --ai-search-key ${searchKey}`;
        const { success } = await Executor.execute(
          installCmd,
          projectPath,
          undefined,
          undefined,
          "will be ignored"
        );
        if (!success) {
          throw new Error("Failed to install packages");
        }
      }

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
        Env.password
      );
      await localDebugTestContext.validateLocalStateForBot();
      if (isRealKey) {
        await validateWelcomeAndReplyBot(page, {
          hasWelcomeMessage: false,
          hasCommandReplyValidation: true,
          botCommand: "Tell me about Contoso Electronics PerksPlus Program",
          expectedWelcomeMessage: ValidationContent.AiChatBotWelcomeInstruction,
          expectedReplyMessage: "$1000",
          timeout: Timeout.longTimeWait,
        });
      } else {
        await validateWelcomeAndReplyBot(page, {
          hasWelcomeMessage: false,
          hasCommandReplyValidation: true,
          botCommand: "helloWorld",
          expectedWelcomeMessage: ValidationContent.AiChatBotWelcomeInstruction,
          expectedReplyMessage: ValidationContent.AiBotErrorMessage,
          timeout: Timeout.longTimeWait,
        });
      }
    }
  );
});

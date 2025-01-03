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
      customCopilotRagType: "custom-copilot-rag-azureAISearch",
    });
    await localDebugTestContext.before();
  });

  afterEach(async function () {
    this.timeout(Timeout.finishTestCase);
    await localDebugTestContext.after(false, true, true);
  });

  it(
    "[auto][JS][Azure OpenAI] Local debug for basic rag bot using azure ai data",
    {
      testPlanCaseId: 28970334,
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

      const embeddingDeploymentName =
        OpenAiKey.azureOpenAiEmbeddingDeploymentName ?? "fake";

      const searchKey = isRealKey ? azSearchHelper.apiKey : "fake";
      const searchEndpoint = isRealKey
        ? azSearchHelper.endpoint
        : "https://test.com";
      editDotEnvFile(envPath, "SECRET_AZURE_OPENAI_API_KEY", azureOpenAiKey);
      editDotEnvFile(envPath, "AZURE_OPENAI_ENDPOINT", azureOpenAiEndpoint);
      editDotEnvFile(
        envPath,
        "AZURE_OPENAI_DEPLOYMENT_NAME",
        azureOpenAiModelDeploymentName
      );
      editDotEnvFile(
        envPath,
        "AZURE_OPENAI_EMBEDDING_DEPLOYMENT_NAME",
        embeddingDeploymentName
      );
      editDotEnvFile(envPath, "SECRET_AZURE_SEARCH_KEY", searchKey);
      editDotEnvFile(envPath, "AZURE_SEARCH_ENDPOINT", searchEndpoint);

      console.log(`
        SECRET_AZURE_OPENAI_API_KEY=${azureOpenAiKey}
        AZURE_OPENAI_ENDPOINT=${azureOpenAiEndpoint}
        AZURE_OPENAI_DEPLOYMENT_NAME=${azureOpenAiModelDeploymentName}
        AZURE_OPENAI_EMBEDDING_DEPLOYMENT_NAME=${embeddingDeploymentName}
        SECRET_AZURE_SEARCH_KEY=${searchKey}
        AZURE_SEARCH_ENDPOINT=${searchEndpoint}
      `);

      // prepare for the npm run indexer:create
      const testToolEnvPath = path.resolve(
        projectPath,
        "env",
        ".env.testtool.user"
      );
      editDotEnvFile(
        testToolEnvPath,
        "SECRET_AZURE_OPENAI_API_KEY",
        azureOpenAiKey
      );
      editDotEnvFile(
        testToolEnvPath,
        "AZURE_OPENAI_ENDPOINT",
        azureOpenAiEndpoint
      );
      editDotEnvFile(
        testToolEnvPath,
        "AZURE_OPENAI_DEPLOYMENT_NAME",
        azureOpenAiModelDeploymentName
      );
      editDotEnvFile(
        testToolEnvPath,
        "AZURE_OPENAI_EMBEDDING_DEPLOYMENT_NAME",
        embeddingDeploymentName
      );
      editDotEnvFile(testToolEnvPath, "SECRET_AZURE_SEARCH_KEY", searchKey);
      editDotEnvFile(testToolEnvPath, "AZURE_SEARCH_ENDPOINT", searchEndpoint);

      // create azure search data
      if (isRealKey) {
        console.log("Start to create azure search data");
        const installCmd = `npm install`;
        const { success } = await Executor.execute(
          installCmd,
          projectPath,
          process.env,
          undefined,
          "npm warn"
        );
        if (!success) {
          throw new Error("Failed to install packages");
        }

        const insertDataCmd = `npm run indexer:create -- ${searchKey} ${azureOpenAiKey}`;
        const { success: insertDataSuccess } = await Executor.execute(
          insertDataCmd,
          projectPath
        );
        if (!insertDataSuccess) {
          throw new Error("Failed to insert data");
        }
      }

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
      if (isRealKey) {
        await validateWelcomeAndReplyBot(page, {
          hasWelcomeMessage: false,
          hasCommandReplyValidation: true,
          botCommand: "Tell me about Contoso Electronics history",
          expectedWelcomeMessage: ValidationContent.AiChatBotWelcomeInstruction,
          expectedReplyMessage: "1985",
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

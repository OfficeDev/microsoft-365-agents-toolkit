// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */
import * as path from "path";
import { VSBrowser } from "vscode-extension-tester";
import {
  createEnvironmentWithPython,
  startDebugging,
  waitForTerminal,
  execCommandIfExist,
  createNewProject,
} from "../../utils/vscodeOperation";
import {
  initPage,
  validateWelcomeAndReplyBot,
} from "../../utils/playwrightOperation";
import { LocalDebugTestContext } from "./localdebugContext";
import {
  RemoteDebugTestContext,
  deployProject,
  provisionProject,
} from "../remotedebug/remotedebugContext";
import {
  Timeout,
  LocalDebugTaskLabel,
  DebugItemSelect,
  ValidationContent,
  LocalDebugTaskLabel2,
} from "../../utils/constants";
import { Env, OpenAiKey } from "../../utils/env";
import { it } from "../../utils/it";
import {
  editDotEnvFile,
  validateFileExist,
  modifyFileContext,
} from "../../utils/commonUtils";
import { Executor } from "../../utils/executor";
import os from "os";

export function happyPathTest(options: {
  lang: "JavaScript" | "TypeScript" | "Python";
  llm: "llm-service-openai" | "llm-service-azure-openai";
  agent: "custom-copilot-agent-new" | "custom-copilot-agent-assistants-api";
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
        debugContent = "local";
      } else {
        debugContent = "remote";
      }

      // ensure workbench is ready
      this.timeout(Timeout.prepareTestCase);
      if (debugContent === "local") {
        localDebugTestContext = new LocalDebugTestContext("aiagent", {
          lang: options.lang,
          customCeopilotAgent: options.agent,
          llmServiceType: options.llm,
        });
        await localDebugTestContext.before();
      } else {
        remoteDebugTestContext = new RemoteDebugTestContext("aiagent");
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
      }] Local debug for AI Agent - ${
        options.agent === "custom-copilot-agent-new"
          ? "Build New"
          : "Build with Assistants API"
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
          `src/${
            options.lang === "JavaScript"
              ? "index.js"
              : options.lang === "TypeScript"
              ? "index.ts"
              : "app.py"
          }`
        );
        // python prepare env
        if (options.lang === "Python") {
          await createEnvironmentWithPython();
        }

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
            "AZURE_OPENAI_MODEL_DEPLOYMENT_NAME",
            azureOpenAiModelDeploymentName
          );

          // js/ts run run assistant:create command
          if (options.lang === "JavaScript" || options.lang === "TypeScript") {
            const creatorFile = path.resolve(
              projectPath,
              "src",
              `creator.${options.lang === "JavaScript" ? "js" : "ts"}`
            );
            modifyFileContext(
              creatorFile,
              'const azureOpenAIEndpoint="";',
              `const azureOpenAIEndpoint="${azureOpenAiEndpoint}";`
            );
            modifyFileContext(
              creatorFile,
              'const azureOpenAIDeploymentName="";',
              `const azureOpenAIDeploymentName="${azureOpenAiModelDeploymentName}";`
            );

            if (
              isRealKey &&
              options.agent === "custom-copilot-agent-assistants-api"
            ) {
              console.log("Start to create azure assistant id");
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

              let insertDataCmd = "";
              if (os.type() === "Windows_NT") {
                insertDataCmd = `npm run assistant:create -- ${azureOpenAiKey}`;
              } else {
                insertDataCmd = `npm run assistant:create -- '${azureOpenAiKey}'`;
              }
              const { success: insertDataSuccess, stdout: log } =
                await Executor.execute(insertDataCmd, projectPath);
              // get assistant id from log string
              const assistantId = log.match(
                /Created a new assistant with an ID of: (.*)/
              )?.[1];
              if (!insertDataSuccess) {
                throw new Error("Failed to create assistant");
              }
              editDotEnvFile(
                envPath,
                "AZURE_OPENAI_ASSISTANT_ID",
                assistantId ?? ""
              );
            } else {
              editDotEnvFile(envPath, "AZURE_OPENAI_ASSISTANT_ID", "fake");
            }
          } else {
            // python run creator.py command
            if (
              isRealKey &&
              options.agent === "custom-copilot-agent-assistants-api"
            ) {
              console.log("Start to create azure assistant id");

              let insertDataCmd = "";
              if (os.type() === "Windows_NT") {
                insertDataCmd = `python src/utils/creator.py --api-key ${azureOpenAiKey}`;
              } else {
                insertDataCmd = `python src/utils/creator.py --api-key '${azureOpenAiKey}'`;
              }

              const { success: insertDataSuccess, stdout: log } =
                await Executor.execute(insertDataCmd, projectPath);
              // get assistant id from log string
              const assistantId = log.match(
                /Created a new assistant with an ID of: (.*)/
              )?.[1];
              if (!insertDataSuccess) {
                throw new Error("Failed to create assistant");
              }
              editDotEnvFile(
                envPath,
                "AZURE_OPENAI_ASSISTANT_ID",
                assistantId ?? ""
              );
            } else {
              editDotEnvFile(envPath, "AZURE_OPENAI_ASSISTANT_ID", "fake");
            }
          }
        } else {
          // openai entrance
          editDotEnvFile(envPath, "SECRET_OPENAI_API_KEY", "fake");
          editDotEnvFile(envPath, "OPENAI_ASSISTANT_ID", "fake");
        }

        await startDebugging(DebugItemSelect.DebugInTeamsUsingChrome);

        await waitForTerminal(LocalDebugTaskLabel.StartLocalTunnel);
        if (options.lang === "Python") {
          await waitForTerminal(
            LocalDebugTaskLabel2.PythonDebugConsole,
            "Running on http://localhost:3978"
          );
        } else {
          await waitForTerminal(LocalDebugTaskLabel.StartBotApp, "Bot Started");
        }

        const teamsAppId = await localDebugTestContext.getTeamsAppId();
        const page = await initPage(
          localDebugTestContext.context!,
          teamsAppId,
          Env.username,
          Env.password,
          { projectPath: projectPath, env: "local" }
        );
        await localDebugTestContext.validateLocalStateForBot();
        if (options.agent === "custom-copilot-agent-new") {
          if (isRealKey) {
            await validateWelcomeAndReplyBot(page, {
              hasCommandReplyValidation: true,
              botCommand: "Remind me to attend the team meeting next Monday",
              expectedReplyMessage:
                "Remind me to attend the team meeting next Monday",
            });
            try {
              await validateWelcomeAndReplyBot(page, {
                hasCommandReplyValidation: true,
                botCommand: "Show all tasks",
                expectedReplyMessage: "current tasks",
                timeout: Timeout.longTimeWait,
              });
            } catch (error) {
              await validateWelcomeAndReplyBot(page, {
                hasCommandReplyValidation: true,
                botCommand: "Show all tasks",
                expectedReplyMessage: ValidationContent.AiBotMeetingMessage,
                timeout: Timeout.longTimeWait,
              });
            }
          } else {
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
        } else {
          if (isRealKey) {
            await validateWelcomeAndReplyBot(page, {
              hasWelcomeMessage: false,
              hasCommandReplyValidation: true,
              botCommand:
                "I need to solve the equation `3x + 11 = 14`. Can you help me?",
              expectedWelcomeMessage:
                ValidationContent.AiAssistantBotWelcomeInstruction,
              expectedReplyMessage: "x = 1",
              timeout: Timeout.longTimeWait,
            });
          } else {
            try {
              await validateWelcomeAndReplyBot(page, {
                hasWelcomeMessage: false,
                hasCommandReplyValidation: true,
                botCommand: "helloWorld",
                expectedWelcomeMessage:
                  ValidationContent.AiAssistantBotWelcomeInstruction,
                expectedReplyMessage: ValidationContent.AiBotErrorMessage,
                timeout: Timeout.longTimeWait,
              });
            } catch (error) {
              await validateWelcomeAndReplyBot(page, {
                hasWelcomeMessage: false,
                hasCommandReplyValidation: true,
                botCommand: "helloWorld",
                expectedWelcomeMessage:
                  ValidationContent.AiAssistantBotWelcomeInstruction,
                expectedReplyMessage: ValidationContent.AiBotErrorMessage2,
                timeout: Timeout.longTimeWait,
              });
            }
          }
        }
      }
    );

    it(
      `[auto][${options.lang}][${
        options.llm === "llm-service-azure-openai" ? "Azure OpenAI" : "OpenAI"
      }] Remote debug for AI Agent - ${
        options.agent === "custom-copilot-agent-new"
          ? "Build New"
          : "Build with Assistants API"
      }`,
      {
        testPlanCaseId: options.testPlanCaseId_dev,
        author: options.author,
      },
      async function () {
        const driver = VSBrowser.instance.driver;
        await createNewProject(
          options.agent === "custom-copilot-agent-new"
            ? "aiagentnew"
            : "aiagentassist",
          appName,
          {
            lang: options.lang,
            aiType:
              options.llm === "llm-service-azure-openai"
                ? "Azure OpenAI"
                : "OpenAI",
          }
        );
        validateFileExist(
          projectPath,
          `src/${
            options.lang === "JavaScript"
              ? "index.js"
              : options.lang === "TypeScript"
              ? "index.ts"
              : "app.py"
          }`
        );
        // python prepare env
        if (options.lang === "Python") {
          await createEnvironmentWithPython();
        }
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
            "AZURE_OPENAI_MODEL_DEPLOYMENT_NAME",
            azureOpenAiModelDeploymentName
          );
        } else {
          // openai entrance
          editDotEnvFile(envPath, "SECRET_OPENAI_API_KEY", "fake");
          editDotEnvFile(envPath, "OPENAI_ASSISTANT_ID", "fake");
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
            "AZURE_OPENAI_MODEL_DEPLOYMENT_NAME",
            azureOpenAiModelDeploymentName
          );
          if (options.lang === "JavaScript" || options.lang === "TypeScript") {
            const creatorFile = path.resolve(
              projectPath,
              "src",
              `creator.${options.lang === "JavaScript" ? "js" : "ts"}`
            );
            modifyFileContext(
              creatorFile,
              'const azureOpenAIEndpoint="";',
              `const azureOpenAIEndpoint="${azureOpenAiEndpoint}";`
            );
            modifyFileContext(
              creatorFile,
              'const azureOpenAIDeploymentName="";',
              `const azureOpenAIDeploymentName="${azureOpenAiModelDeploymentName}";`
            );

            if (
              isRealKey &&
              options.agent === "custom-copilot-agent-assistants-api"
            ) {
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

              let insertDataCmd = "";
              if (os.type() === "Windows_NT") {
                insertDataCmd = `npm run assistant:create -- ${azureOpenAiKey}`;
              } else {
                insertDataCmd = `npm run assistant:create -- '${azureOpenAiKey}'`;
              }
              const { success: insertDataSuccess, stdout: log } =
                await Executor.execute(insertDataCmd, projectPath);
              // get assistant id from log string
              const assistantId = log.match(
                /Created a new assistant with an ID of: (.*)/
              )?.[1];
              if (!insertDataSuccess) {
                throw new Error("Failed to create assistant");
              }
              editDotEnvFile(
                envPath,
                "AZURE_OPENAI_ASSISTANT_ID",
                assistantId ?? ""
              );
            } else {
              editDotEnvFile(envPath, "AZURE_OPENAI_ASSISTANT_ID", "fake");
            }
          } else {
            if (
              isRealKey &&
              options.agent === "custom-copilot-agent-assistants-api"
            ) {
              console.log("Start to create azure assistant id");

              let insertDataCmd = "";
              if (os.type() === "Windows_NT") {
                insertDataCmd = `python src/utils/creator.py --api-key ${azureOpenAiKey}`;
              } else {
                insertDataCmd = `python src/utils/creator.py --api-key '${azureOpenAiKey}'`;
              }
              const { success: insertDataSuccess, stdout: log } =
                await Executor.execute(insertDataCmd, projectPath);
              // get assistant id from log string
              const assistantId = log.match(
                /Created a new assistant with an ID of: (.*)/
              )?.[1];
              if (!insertDataSuccess) {
                throw new Error("Failed to create assistant");
              }
              editDotEnvFile(
                envPath,
                "AZURE_OPENAI_ASSISTANT_ID",
                assistantId ?? ""
              );
            } else {
              editDotEnvFile(envPath, "AZURE_OPENAI_ASSISTANT_ID", "fake");
            }
          }
        }

        await provisionProject(appName, projectPath);
        await deployProject(projectPath, Timeout.botDeploy);
        // [known issue] python remote need deploy twice
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

        if (options.agent === "custom-copilot-agent-new") {
          if (isRealKey) {
            await validateWelcomeAndReplyBot(page, {
              hasCommandReplyValidation: true,
              botCommand: "Remind me to attend the team meeting next Monday",
              expectedReplyMessage:
                "Remind me to attend the team meeting next Monday",
            });
            try {
              await validateWelcomeAndReplyBot(page, {
                hasCommandReplyValidation: true,
                botCommand: "Show all tasks",
                expectedReplyMessage: "current tasks",
                timeout: Timeout.longTimeWait,
              });
            } catch (error) {
              await validateWelcomeAndReplyBot(page, {
                hasCommandReplyValidation: true,
                botCommand: "Show all tasks",
                expectedReplyMessage: ValidationContent.AiBotMeetingMessage,
                timeout: Timeout.longTimeWait,
              });
            }
          } else {
            try {
              await validateWelcomeAndReplyBot(page, {
                hasWelcomeMessage: false,
                hasCommandReplyValidation: true,
                botCommand: "helloWorld",
                expectedWelcomeMessage:
                  ValidationContent.AiAssistantBotWelcomeInstruction,
                expectedReplyMessage: ValidationContent.AiBotErrorMessage,
                timeout: Timeout.longTimeWait,
              });
            } catch (error) {
              await validateWelcomeAndReplyBot(page, {
                hasWelcomeMessage: false,
                hasCommandReplyValidation: true,
                botCommand: "helloWorld",
                expectedWelcomeMessage:
                  ValidationContent.AiAssistantBotWelcomeInstruction,
                expectedReplyMessage: ValidationContent.AiBotErrorMessage2,
                timeout: Timeout.longTimeWait,
              });
            }
          }
        } else {
          if (isRealKey) {
            await validateWelcomeAndReplyBot(page, {
              hasWelcomeMessage: false,
              hasCommandReplyValidation: true,
              botCommand:
                "I need to solve the equation `3x + 11 = 14`. Can you help me?",
              expectedWelcomeMessage:
                ValidationContent.AiAssistantBotWelcomeInstruction,
              expectedReplyMessage: "x = 1",
              timeout: Timeout.longTimeWait,
            });
          } else {
            try {
              await validateWelcomeAndReplyBot(page, {
                hasWelcomeMessage: false,
                hasCommandReplyValidation: true,
                botCommand: "helloWorld",
                expectedWelcomeMessage:
                  ValidationContent.AiAssistantBotWelcomeInstruction,
                expectedReplyMessage: ValidationContent.AiBotErrorMessage,
                timeout: Timeout.longTimeWait,
              });
            } catch (error) {
              await validateWelcomeAndReplyBot(page, {
                hasWelcomeMessage: false,
                hasCommandReplyValidation: true,
                botCommand: "helloWorld",
                expectedWelcomeMessage:
                  ValidationContent.AiAssistantBotWelcomeInstruction,
                expectedReplyMessage: ValidationContent.AiBotErrorMessage2,
                timeout: Timeout.longTimeWait,
              });
            }
          }
        }
      }
    );
  });
}

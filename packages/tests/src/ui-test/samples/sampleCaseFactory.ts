// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import {
  Timeout,
  TemplateProject,
  sampleProjectMap,
  LocalDebugTaskLabel,
  LocalDebugTaskResult,
  LocalDebugError,
} from "../../utils/constants";
import { waitForTerminal, stopDebugging } from "../../utils/vscodeOperation";
import {
  debugInitMap,
  initPage,
  reopenPage,
} from "../../utils/playwrightOperation";
import { Env } from "../../utils/env";
import { SampledebugContext } from "./sampledebugContext";
import { it } from "../../utils/it";
import { VSBrowser } from "vscode-extension-tester";
import { getScreenshotName } from "../../utils/nameUtil";
import { AzSqlHelper } from "../../utils/azureCliHelper";
import { expect } from "chai";
import { Page } from "playwright";
import fs from "fs-extra";
import path from "path";
import { Executor } from "../../utils/executor";
import { ChildProcess, ChildProcessWithoutNullStreams } from "child_process";
import { initDebugPort } from "../../utils/commonUtils";
import { CliHelper } from "../cliHelper";

const debugMap: Record<LocalDebugTaskLabel, () => Promise<void>> = {
  [LocalDebugTaskLabel.StartFrontend]: async () => {
    await waitForTerminal(
      LocalDebugTaskLabel.StartFrontend,
      LocalDebugTaskResult.FrontendSuccess
    );
  },
  [LocalDebugTaskLabel.StartBackend]: async () => {
    await waitForTerminal(
      LocalDebugTaskLabel.StartBackend,
      LocalDebugTaskResult.BotAppSuccess
    );
  },
  [LocalDebugTaskLabel.WatchBackend]: async () => {
    await waitForTerminal(
      LocalDebugTaskLabel.WatchBackend,
      LocalDebugTaskResult.CompiledSuccess
    );
  },
  [LocalDebugTaskLabel.StartLocalTunnel]: async () => {
    await waitForTerminal(
      LocalDebugTaskLabel.StartLocalTunnel,
      LocalDebugTaskResult.StartSuccess
    );
  },
  [LocalDebugTaskLabel.Azurite]: async () => {
    await waitForTerminal(
      LocalDebugTaskLabel.Azurite,
      LocalDebugTaskResult.AzuriteSuccess
    );
  },
  [LocalDebugTaskLabel.Compile]: async () => {
    await waitForTerminal(
      LocalDebugTaskLabel.Compile,
      LocalDebugTaskResult.CompiledSuccess
    );
  },
  [LocalDebugTaskLabel.StartBotApp]: async () => {
    await waitForTerminal(
      LocalDebugTaskLabel.StartBotApp,
      LocalDebugTaskResult.BotAppSuccess
    );
  },
  [LocalDebugTaskLabel.StartBot]: async () => Promise.resolve(),
  [LocalDebugTaskLabel.StartWebhook]: async () => {
    await waitForTerminal(
      LocalDebugTaskLabel.StartWebhook,
      LocalDebugTaskResult.DebuggerAttached
    );
  },
  [LocalDebugTaskLabel.InstallNpmPackages]: async () => Promise.resolve(),
  [LocalDebugTaskLabel.ApiNpmInstall]: async () => Promise.resolve(),
  [LocalDebugTaskLabel.BotNpmInstall]: async () => Promise.resolve(),
  [LocalDebugTaskLabel.TabsNpmInstall]: async () => Promise.resolve(),
  [LocalDebugTaskLabel.SpfxNpmInstall]: async () => Promise.resolve(),
  [LocalDebugTaskLabel.GulpServe]: async () => {
    await waitForTerminal(
      LocalDebugTaskLabel.GulpServe,
      LocalDebugTaskResult.GulpServeSuccess
    );
  },
  [LocalDebugTaskLabel.StartWebServer]: async () => {
    await waitForTerminal(
      LocalDebugTaskLabel.StartWebServer,
      LocalDebugTaskResult.WebServerSuccess
    );
  },
  [LocalDebugTaskLabel.DockerRun]: async () => {
    await waitForTerminal(
      LocalDebugTaskLabel.DockerRun,
      LocalDebugTaskResult.DockerFinish
    );
  },
  [LocalDebugTaskLabel.DockerTask]: async () => {
    await waitForTerminal(
      LocalDebugTaskLabel.DockerTask,
      LocalDebugTaskResult.DockerFinish
    );
  },
  [LocalDebugTaskLabel.EnsureDevTunnnel]: async () => {
    await waitForTerminal(
      LocalDebugTaskLabel.EnsureDevTunnnel,
      LocalDebugTaskResult.DevtunnelSuccess
    );
  },
  [LocalDebugTaskLabel.RunWatch]: async () => {
    await waitForTerminal(
      LocalDebugTaskLabel.RunWatch,
      LocalDebugTaskResult.CompiledSuccess
    );
  },
  [LocalDebugTaskLabel.FuncStart]: async () => {
    await waitForTerminal(LocalDebugTaskLabel.FuncStart);
  },
};

export abstract class CaseFactory {
  public sampleName: TemplateProject;
  public testPlanCaseId: number;
  public author: string;
  public env: "local" | "dev";
  public validate: LocalDebugTaskLabel[];
  public options?: {
    teamsAppName?: string;
    dashboardFlag?: boolean;
    type?: string;
    testRootFolder?: string;
    includeFunction?: boolean;
    npmName?: string;
    skipInit?: boolean;
    skipValidation?: boolean;
    skipDebug?: boolean;
    debug?: "cli" | "ttk";
    botFlag?: boolean;
    repoPath?: string;
    container?: boolean;
    dockerFolder?: string;
    skipDeploy?: boolean;
  };

  public constructor(
    sampleName: TemplateProject,
    testPlanCaseId: number,
    author: string,
    env: "local" | "dev",
    validate: LocalDebugTaskLabel[] = [],
    options: {
      teamsAppName?: string;
      dashboardFlag?: boolean;
      type?: string;
      testRootFolder?: string;
      includeFunction?: boolean;
      npmName?: string;
      skipInit?: boolean;
      skipValidation?: boolean;
      skipDebug?: boolean;
      debug?: "cli" | "ttk";
      botFlag?: boolean;
      repoPath?: string;
      container?: boolean;
      dockerFolder?: string;
      skipDeploy?: boolean;
    } = {}
  ) {
    this.sampleName = sampleName;
    this.testPlanCaseId = testPlanCaseId;
    this.author = author;
    this.env = env;
    this.validate = validate;
    this.options = options;
  }

  public onBefore(
    sampledebugContext: SampledebugContext,
    env: "local" | "dev",
    azSqlHelper?: AzSqlHelper
  ): Promise<AzSqlHelper | undefined> {
    return Promise.resolve(undefined);
  }

  public async onAfter(
    sampledebugContext: SampledebugContext,
    env: "local" | "dev"
  ): Promise<void> {
    const envMap: Record<
      "local" | "dev",
      (options?: { rgName: string }) => Promise<void>
    > = {
      local: async () => await sampledebugContext.after(),
      dev: async (options?: { rgName: string }) =>
        await sampledebugContext.sampleAfter(options?.rgName ?? ""),
    };
    await envMap[env]({ rgName: `${sampledebugContext.appName}-dev-rg` });
  }

  public async onAfterCreate(
    sampledebugContext: SampledebugContext,
    env: "local" | "dev",
    azSqlHelper?: AzSqlHelper
  ): Promise<void> {
    return Promise.resolve();
  }

  public async onBeforeBrowerStart(
    sampledebugContext: SampledebugContext,
    env: "local" | "dev",
    azSqlHelper?: AzSqlHelper
  ): Promise<void> {
    return Promise.resolve();
  }

  public async onInitPage(
    sampledebugContext: SampledebugContext,
    teamsAppId: string,
    options?: {
      teamsAppName: string;
      includeFunction: boolean;
      npmName: string;
      dashboardFlag: boolean;
      type: string;
      env: "local" | "dev";
    }
  ): Promise<Page> {
    return await initPage(
      sampledebugContext.context!,
      teamsAppId,
      Env.username,
      Env.password,
      {
        projectPath: sampledebugContext.projectPath,
        env: options?.env,
        dashboardFlag: options?.dashboardFlag,
      }
    );
  }

  public async onReopenPage(
    sampledebugContext: SampledebugContext,
    teamsAppId: string,
    options?: {
      teamsAppName: string;
      includeFunction: boolean;
      npmName: string;
      dashboardFlag: boolean;
      type: string;
      env: "local" | "dev";
    }
  ): Promise<Page> {
    return await reopenPage(
      sampledebugContext.context!,
      teamsAppId,
      Env.username,
      Env.password,
      {
        projectPath: sampledebugContext.projectPath,
        env: options?.env,
        dashboardFlag: options?.dashboardFlag,
      }
    );
  }

  public async onValidate(
    page: Page,
    options?: {
      context: SampledebugContext;
      displayName: string;
      includeFunction: boolean;
      npmName: string;
      env: "local" | "dev";
    }
  ): Promise<void> {
    Promise.resolve();
  }

  public async onCliValidate(
    page: Page,
    options?: {
      context: SampledebugContext;
      displayName: string;
      includeFunction: boolean;
      npmName: string;
      env: "local" | "dev";
    }
  ): Promise<void> {
    Promise.resolve();
  }

  public test(): void {
    const {
      sampleName,
      testPlanCaseId,
      author,
      env,
      validate,
      options,
      onBefore,
      onAfterCreate,
      onAfter,
      onBeforeBrowerStart,
      onInitPage,
      onValidate,
      onReopenPage,
      onCliValidate,
    } = this;
    describe("Sample Tests", function () {
      this.timeout(Timeout.testAzureCase);
      let sampledebugContext: SampledebugContext;
      let azSqlHelper: AzSqlHelper | undefined;
      let devtunnelProcess: ChildProcessWithoutNullStreams;
      let debugProcess: ChildProcess;
      let dockerProcess: ChildProcessWithoutNullStreams;
      let successFlag = true;
      let envContent = "";
      let botFlag = false;
      let envFile = "";
      let errorMessage = "";

      beforeEach(async function () {
        // ensure workbench is ready
        this.timeout(Timeout.prepareTestCase);
        sampledebugContext = new SampledebugContext(
          sampleName,
          sampleProjectMap[sampleName],
          options?.testRootFolder ?? "./resource",
          options?.repoPath ?? "./resource"
        );
        await sampledebugContext.before();
        // use before middleware to process typical sample
        azSqlHelper = await onBefore(sampledebugContext, env, azSqlHelper);
      });

      afterEach(async function () {
        this.timeout(Timeout.finishTestCase);
        await onAfter(sampledebugContext, env);
        setTimeout(() => {
          if (successFlag) process.exit(0);
          else process.exit(1);
        }, 30000);
      });

      it(
        `[auto] local debug for Sample ${sampleName}`,
        {
          testPlanCaseId,
          author,
        },
        async function () {
          try {
            // create project
            await sampledebugContext.openResourceFolder();
            // update manifest app name
            await sampledebugContext.updateManifestAppName();
            // use 1st middleware to process typical sample
            await onAfterCreate(sampledebugContext, env, azSqlHelper);

            try {
              envFile = path.resolve(
                sampledebugContext.projectPath,
                "env",
                ".env.local"
              );
              envContent = fs.readFileSync(envFile, "utf-8");
              // if bot project setup devtunnel
              botFlag = envContent.includes("BOT_DOMAIN");
            } catch (error) {
              console.log("read file error", error);
            }

            if (options?.skipDebug) {
              console.log("skip ui skipDebug...");
              console.log("debug finish!");
              return;
            }

            // cli preview
            if (options?.debug === "cli") {
              console.log("======= debug with cli ========");
              // start local tunnel
              if (options.botFlag || botFlag) {
                const tunnel = Executor.debugBotFunctionPreparation(
                  sampledebugContext.projectPath
                );
                devtunnelProcess = tunnel.devtunnelProcess;
              }
              await new Promise((resolve) => setTimeout(resolve, 60 * 1000));
              const { success: provisionSuccess } = await Executor.provision(
                sampledebugContext.projectPath,
                "local",
                true,
                "DeprecationWarning"
              );
              expect(provisionSuccess).to.be.true;
              if (!options.container) {
                const { success: deploySuccess } = await Executor.deploy(
                  sampledebugContext.projectPath,
                  "local"
                );
                expect(deploySuccess).to.be.true;
              } else {
                await CliHelper.dockerBuild(
                  sampledebugContext.projectPath,
                  options.dockerFolder || ""
                );

                dockerProcess = await CliHelper.dockerRun(
                  sampledebugContext.projectPath,
                  options.dockerFolder || ""
                );
              }
              const teamsAppId = await sampledebugContext.getTeamsAppId(env);
              expect(teamsAppId).to.not.be.empty;

              debugProcess = Executor.debugProject(
                sampledebugContext.projectPath,
                "local",
                true,
                process.env,
                (data) => {
                  if (data) {
                    console.log(data);
                  }
                },
                (error) => {
                  const errorMsg = error.toString();
                  console.log("[error log]", errorMsg);
                  if (
                    // skip warning messages
                    errorMsg.includes(LocalDebugError.WarningError)
                  ) {
                    console.log("[skip error] ", error);
                  } else {
                    successFlag = false;
                    expect.fail(errorMsg);
                  }
                },
                options.container
              );
              await new Promise((resolve) =>
                setTimeout(resolve, 2 * 60 * 1000)
              );

              // if no skip init step
              if (!options?.skipInit) {
                // init
                const page = await onInitPage(sampledebugContext, teamsAppId, {
                  includeFunction: options?.includeFunction ?? false,
                  npmName: options?.npmName ?? "",
                  dashboardFlag: options?.dashboardFlag ?? false,
                  type: options?.type ?? "",
                  teamsAppName: options?.teamsAppName ?? "",
                  env: env,
                });

                // if no skip vaildation
                if (!options?.skipValidation) {
                  await onCliValidate(page, {
                    context: sampledebugContext,
                    displayName: Env.displayName,
                    includeFunction: options?.includeFunction ?? false,
                    npmName: options?.npmName ?? "",
                    env: env,
                  });
                } else {
                  console.log("skip ui skipValidation...");
                  console.log("debug finish!");
                }
              } else {
                console.log("skip ui skipInit...");
                console.log("debug finish!");
              }
              // kill process
              await Executor.closeProcess(debugProcess);
              if (botFlag) await Executor.closeProcess(devtunnelProcess);
              if (dockerProcess) {
                await Executor.closeProcess(dockerProcess);
                await CliHelper.stopAllDocker();
              }
              await initDebugPort();
            }

            // ttk debug
            // local debug with ttk
            console.log("======= debug with ttk ========");
            await debugInitMap[sampleName]();
            for (const label of validate) {
              try {
                await debugMap[label]();
              } catch (error) {
                const errorMsg = error.toString();
                if (
                  // skip can't find element
                  errorMsg.includes(
                    LocalDebugError.ElementNotInteractableError
                  ) ||
                  // skip timeout
                  errorMsg.includes(LocalDebugError.TimeoutError)
                ) {
                  console.log("[skip error] ", error);
                } else {
                  expect.fail(errorMsg);
                }
              }
            }

            // if no skip init step
            if (!options?.skipInit) {
              const teamsAppId = await sampledebugContext.getTeamsAppId(env);
              expect(teamsAppId).to.not.be.empty;
              // use 2nd middleware to process typical sample
              await onBeforeBrowerStart(sampledebugContext, env, azSqlHelper);
              // init
              let page: Page;
              if (options?.debug === "cli") {
                page = await onReopenPage(sampledebugContext, teamsAppId, {
                  includeFunction: options?.includeFunction ?? false,
                  npmName: options?.npmName ?? "",
                  dashboardFlag: options?.dashboardFlag ?? false,
                  type: options?.type ?? "",
                  teamsAppName: options?.teamsAppName ?? "",
                  env: env,
                });
              } else {
                page = await onInitPage(sampledebugContext, teamsAppId, {
                  includeFunction: options?.includeFunction ?? false,
                  npmName: options?.npmName ?? "",
                  dashboardFlag: options?.dashboardFlag ?? false,
                  type: options?.type ?? "",
                  teamsAppName: options?.teamsAppName ?? "",
                  env: env,
                });
              }

              // if no skip vaildation
              if (!options?.skipValidation) {
                await onValidate(page, {
                  context: sampledebugContext,
                  displayName: Env.displayName,
                  includeFunction: options?.includeFunction ?? false,
                  npmName: options?.npmName ?? "",
                  env: env,
                });
              } else {
                console.log("skip ui skipValidation...");
                console.log("debug finish!");
              }
              await stopDebugging();
            } else {
              console.log("skip ui skipInit...");
              console.log("debug finish!");
            }
          } catch (error) {
            successFlag = false;
            errorMessage = "[Error]: " + error;
            await VSBrowser.instance.takeScreenshot(getScreenshotName("error"));
            await VSBrowser.instance.driver.sleep(
              Timeout.playwrightDefaultTimeout
            );
          }

          expect(successFlag, errorMessage).to.true;
          console.log("debug finish!");
        }
      );

      it(
        `[auto] remote debug for Sample ${sampleName}`,
        {
          testPlanCaseId,
          author,
        },
        async function () {
          try {
            // create project
            await sampledebugContext.openResourceFolder();
            // update manifest app name
            await sampledebugContext.updateManifestAppName();
            // use 1st middleware to process typical sample
            await onAfterCreate(sampledebugContext, env, azSqlHelper);

            try {
              envFile = path.resolve(
                sampledebugContext.projectPath,
                "env",
                ".env.local"
              );
              envContent = fs.readFileSync(envFile, "utf-8");
              // if bot project setup devtunnel
              botFlag = envContent.includes("BOT_DOMAIN");
            } catch (error) {
              console.log("read file error", error);
            }

            if (options?.skipDebug) {
              console.log("skip ui skipDebug...");
              console.log("debug finish!");
              return;
            }

            // ttk debug
            await sampledebugContext.provisionProject(
              sampledebugContext.appName,
              sampledebugContext.projectPath,
              undefined,
              undefined,
              undefined,
              undefined,
              undefined,
              "DeprecationWarning"
            );
            if (options?.container) {
              await Executor.login();
            }
            if (!options?.skipDeploy) {
              await sampledebugContext.deployProject(
                sampledebugContext.projectPath,
                Timeout.botDeploy
              );
            }

            // if no skip init step
            if (!options?.skipInit) {
              const teamsAppId = await sampledebugContext.getTeamsAppId(env);
              expect(teamsAppId).to.not.be.empty;
              // use 2nd middleware to process typical sample
              await onBeforeBrowerStart(sampledebugContext, env, azSqlHelper);
              // init
              const page = await onInitPage(sampledebugContext, teamsAppId, {
                includeFunction: options?.includeFunction ?? false,
                npmName: options?.npmName ?? "",
                dashboardFlag: options?.dashboardFlag ?? false,
                type: options?.type ?? "",
                teamsAppName: options?.teamsAppName ?? "",
                env: env,
              });

              // if no skip vaildation
              if (!options?.skipValidation) {
                await onValidate(page, {
                  context: sampledebugContext,
                  displayName: Env.displayName,
                  includeFunction: options?.includeFunction ?? false,
                  npmName: options?.npmName ?? "",
                  env: env,
                });
              } else {
                console.log("skip ui skipValidation...");
                console.log("debug finish!");
              }
              await stopDebugging();
            } else {
              console.log("skip ui skipInit...");
              console.log("debug finish!");
            }
          } catch (error) {
            successFlag = false;
            errorMessage = "[Error]: " + error;
            await VSBrowser.instance.takeScreenshot(getScreenshotName("error"));
            await VSBrowser.instance.driver.sleep(
              Timeout.playwrightDefaultTimeout
            );
          }

          expect(successFlag, errorMessage).to.true;
          console.log("debug finish!");
        }
      );
    });
  }
}

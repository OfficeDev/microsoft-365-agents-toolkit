// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Xiaofu Huang <xiaofu.huang@microsoft.com>
 */
import * as path from "path";
import { startDebugging, waitForTerminal } from "../../utils/vscodeOperation";
import {
  initPage,
  reopenPage,
  validateEchoBot,
} from "../../utils/playwrightOperation";
import { LocalDebugTestContext } from "./localdebugContext";
import {
  Timeout,
  LocalDebugTaskLabel,
  DebugItemSelect,
} from "../../utils/constants";
import { Env } from "../../utils/env";
import { it } from "../../utils/it";
import { validateFileExist } from "../../utils/commonUtils";
import { ChildProcess, ChildProcessWithoutNullStreams } from "child_process";
import { Executor } from "../../utils/executor";
import { expect } from "chai";
import { VSBrowser } from "vscode-extension-tester";
import { getScreenshotName } from "../../utils/nameUtil";
import os from "os";
import { initDebugPort } from "../../utils/commonUtils";

describe("Local Debug Tests", function () {
  this.timeout(Timeout.testAzureCase);
  let localDebugTestContext: LocalDebugTestContext;
  let devtunnelProcess: ChildProcessWithoutNullStreams | null;
  let debugProcess: ChildProcess | null;
  let successFlag = true;
  let errorMessage = "";

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    localDebugTestContext = new LocalDebugTestContext("bot", {
      lang: "typescript",
    });
    await localDebugTestContext.before();
  });

  after(async function () {
    this.timeout(Timeout.finishTestCase);
    await localDebugTestContext.after(false, true);
    setTimeout(() => {
      if (os.type() === "Windows_NT") {
        if (successFlag) process.exit(0);
        else process.exit(1);
      }
    }, 30000);
  });

  it(
    "[auto] [Typescript] Local Debug for bot project",
    {
      testPlanCaseId: 9729308,
      author: "xiaofu.huang@microsoft.com",
    },
    async function () {
      try {
        const projectPath = path.resolve(
          localDebugTestContext.testRootFolder,
          localDebugTestContext.appName
        );
        validateFileExist(projectPath, "index.ts");

        // local debug
        console.log("======= debug with ttk ========");
        await startDebugging(DebugItemSelect.DebugInTeamsUsingChrome);
        await waitForTerminal(LocalDebugTaskLabel.StartLocalTunnel);
        await waitForTerminal(LocalDebugTaskLabel.StartBotApp, "Bot Started");

        const teamsAppId = await localDebugTestContext.getTeamsAppId();
        expect(teamsAppId).to.not.be.empty;
        {
          const page = await initPage(
            localDebugTestContext.context!,
            teamsAppId,
            Env.username,
            Env.password,
            { projectPath: projectPath, env: "local" }
          );
          await localDebugTestContext.validateLocalStateForBot();
          await validateEchoBot(page);
        }

        // cli preview
        const res = await Executor.cliPreview(projectPath, true);
        devtunnelProcess = res.devtunnelProcess;
        debugProcess = res.debugProcess;
        {
          const page = await reopenPage(
            localDebugTestContext.context!,
            teamsAppId,
            Env.username,
            Env.password,
            { projectPath: projectPath, env: "local" }
          );
          await localDebugTestContext.validateLocalStateForBot();
          await validateEchoBot(page);
        }
      } catch (error) {
        successFlag = false;
        errorMessage = "[Error]: " + error;
        await VSBrowser.instance.takeScreenshot(getScreenshotName("error"));
        await VSBrowser.instance.driver.sleep(Timeout.playwrightDefaultTimeout);
      }

      // kill process
      await Executor.closeProcess(debugProcess);
      await Executor.closeProcess(devtunnelProcess);
      await initDebugPort();

      expect(successFlag, errorMessage).to.true;
      console.log("debug finish!");
    }
  );
});

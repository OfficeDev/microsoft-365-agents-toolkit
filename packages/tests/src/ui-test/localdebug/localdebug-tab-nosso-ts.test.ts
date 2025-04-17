// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Xiaofu Huang <xiaofu.huang@microsoft.com>
 */
import * as path from "path";
import { startDebugging, waitForTerminal } from "../../utils/vscodeOperation";
import {
  initPage,
  validateBasicTab,
  reopenPage,
} from "../../utils/playwrightOperation";
import { LocalDebugTestContext } from "./localdebugContext";
import {
  Timeout,
  LocalDebugTaskLabel,
  ValidationContent,
  DebugItemSelect,
  LocalDebugTaskResult,
  Lang,
} from "../../utils/constants";
import { Env } from "../../utils/env";
import { it } from "../../utils/it";
import { validateFileExist } from "../../utils/commonUtils";
import { ChildProcess, ChildProcessWithoutNullStreams } from "child_process";
import { Executor } from "../../utils/executor";
import { expect } from "chai";
import { VSBrowser } from "vscode-extension-tester";
import { getScreenshotName } from "../../utils/nameUtil";
import { initDebugPort } from "../../utils/commonUtils";
import os from "os";

describe("Local Debug Tests", function () {
  this.timeout(Timeout.testAzureCase);
  let localDebugTestContext: LocalDebugTestContext;
  let debugProcess: ChildProcess | null;
  let successFlag = true;
  let errorMessage = "";

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    localDebugTestContext = new LocalDebugTestContext("tabnsso", {
      lang: Lang.TS,
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
    "[auto] [Typescript] Local Debug for Tab project without sso",
    {
      testPlanCaseId: 15276946,
      author: "xiaofu.huang@microsoft.com",
    },
    async () => {
      try {
        const projectPath = path.resolve(
          localDebugTestContext.testRootFolder,
          localDebugTestContext.appName
        );
        validateFileExist(projectPath, "src/app.ts");

        console.log("======= debug with ttk ========");
        await startDebugging(DebugItemSelect.DebugInTeamsUsingChrome);
        await waitForTerminal(
          LocalDebugTaskLabel.StartApplication,
          LocalDebugTaskResult.FrontendStarted
        );

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
          await validateBasicTab(page, ValidationContent.Tab);
        }

        // cli preview
        const res = await Executor.cliPreview(projectPath, false);
        debugProcess = res.debugProcess;
        {
          const page = await reopenPage(
            localDebugTestContext.context!,
            teamsAppId,
            Env.username,
            Env.password,
            { projectPath: projectPath, env: "local" }
          );
          await validateBasicTab(page, ValidationContent.Tab);
        }
      } catch (error) {
        successFlag = false;
        errorMessage = "[Error]: " + error;
        await VSBrowser.instance.takeScreenshot(getScreenshotName("error"));
        await VSBrowser.instance.driver.sleep(Timeout.playwrightDefaultTimeout);
      }

      // kill process
      await Executor.closeProcess(debugProcess);
      await initDebugPort();

      expect(successFlag, errorMessage).to.true;
      console.log("debug finish!");
    }
  );
});

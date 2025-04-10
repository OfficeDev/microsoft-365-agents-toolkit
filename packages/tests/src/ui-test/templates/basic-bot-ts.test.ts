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
import { LocalDebugTestContext } from "../localdebug/localdebugContext";
import {
  Timeout,
  LocalDebugTaskLabel,
  DebugItemSelect,
  LocalDebugTaskInfo,
} from "../../utils/constants";
import { Env } from "../../utils/env";
import { it } from "../../utils/it";
import { validateFileExist } from "../../utils/commonUtils";
import { ChildProcess, ChildProcessWithoutNullStreams } from "child_process";
import { Executor } from "../../utils/executor";
import { expect } from "chai";
import { ModalDialog, VSBrowser } from "vscode-extension-tester";
import { getScreenshotName } from "../../utils/nameUtil";
import { initDebugPort } from "../../utils/commonUtils";
import {
  RemoteDebugTestContext,
  provisionProject,
  deployProject,
} from "../remotedebug/remotedebugContext";
import {
  execCommandIfExist,
  createNewProject,
} from "../../utils/vscodeOperation";

describe("Local Debug Tests", function () {
  this.timeout(Timeout.localAndRemoteTestCase);
  let devtunnelProcess: ChildProcessWithoutNullStreams | null;
  let debugProcess: ChildProcess | null;
  let successFlagForLocal = false;
  let successFlagForRemote = false;

  after(async function () {
    this.timeout(Timeout.finishTestCase);
    setTimeout(() => {
      if (successFlagForLocal && successFlagForRemote) process.exit(0);
      else process.exit(1);
    }, 30000);
  });

  it(
    "[auto] [Typescript] Local Debug for bot project",
    {
      testPlanCaseId: 9729308,
      author: "xiaofu.huang@microsoft.com",
    },
    async function () {
      let errorMessage = "";
      const localDebugTestContext = new LocalDebugTestContext("bot", {
        lang: "typescript",
      });
      await localDebugTestContext.before();
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
        await waitForTerminal(
          LocalDebugTaskLabel.StartBotApp,
          LocalDebugTaskInfo.AppListening
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
          successFlagForLocal = true;
        }
      } catch (error) {
        errorMessage = "[Error]: " + error;
        console.log(errorMessage);
        await VSBrowser.instance.takeScreenshot(getScreenshotName("error"));
        await VSBrowser.instance.driver.sleep(Timeout.playwrightDefaultTimeout);
      }
      // kill process
      await Executor.closeProcess(debugProcess);
      await Executor.closeProcess(devtunnelProcess);
      await initDebugPort();

      console.log("debug finish!");
      await localDebugTestContext.after(false, true);
      try {
        //Close the folder and cleanup local sample project
        await execCommandIfExist(
          "Workspaces: Close Workspace",
          Timeout.webView
        );
      } catch {
        const dialog = new ModalDialog();
        console.log(`Click "Cancel" button if it exists`);
        await dialog.pushButton("Cancel");
        console.log(`Clicked button "Cancel"`);
        await execCommandIfExist(
          "Workspaces: Close Workspace",
          Timeout.webView
        );
      }
      expect(successFlagForLocal, errorMessage).to.true;
    }
  );

  it(
    "[auto] Remote debug for bot typescript project Tests",
    {
      testPlanCaseId: 14134645,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      const remoteDebugTestContext = new RemoteDebugTestContext("bot");
      const testRootFolder = remoteDebugTestContext.testRootFolder;
      const appName = remoteDebugTestContext.appName;
      const appNameCopySuffix = "copy";
      const newAppFolderName = appName + appNameCopySuffix;
      const projectPath = path.resolve(testRootFolder, newAppFolderName);
      await remoteDebugTestContext.before();
      const driver = VSBrowser.instance.driver;
      await createNewProject("bot", appName, { lang: "TypeScript" });
      await provisionProject(appName, projectPath);
      try {
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
        await validateEchoBot(page);
        successFlagForRemote = true;
      } catch (error) {
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
        throw new Error("[Error]: " + error);
      }
    }
  );
});

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Xiaofu Huang <xiaofu.huang@microsoft.com>
 */
import * as path from "path";
import { startDebugging, waitForTerminal } from "../../utils/vscodeOperation";
import { initPage, validateEchoBot } from "../../utils/playwrightOperation";
import { LocalDebugTestContext } from "../localdebug/localdebugContext";
import {
  Timeout,
  LocalDebugTaskLabel,
  DebugItemSelect,
  LocalDebugTaskInfo,
  Lang,
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
import { log } from "console";

describe("Local Debug Tests", function () {
  this.timeout(Timeout.localAndRemoteTestCase);
  let devtunnelProcess: ChildProcessWithoutNullStreams | null;
  let debugProcess: ChildProcess | null;
  const successFlagForLocal = true;
  let successFlagForRemote = false;

  after(async function () {
    this.timeout(Timeout.finishTestCase);
    setTimeout(() => {
      if (successFlagForLocal && successFlagForRemote) process.exit(0);
      else process.exit(1);
    }, 30000);
  });

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
      await createNewProject("bot", appName, { lang: Lang.TS });
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
          {
            projectPath: projectPath,
            env: "dev",
            teamsAppName: appName,
            searchApp: false,
          }
        );
        await driver.sleep(Timeout.longTimeWait);
        await validateEchoBot(page, { botCommand: "Hi" });
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

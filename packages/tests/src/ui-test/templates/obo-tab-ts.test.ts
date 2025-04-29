// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Kuojian Lu <kuojianlu@microsoft.com>
 */
import * as path from "path";
import {
  clearNotifications,
  startDebugging,
  waitForTerminal,
} from "../../utils/vscodeOperation";
import {
  initPage,
  validateReactOutlookTab,
  validateReactTab,
} from "../../utils/playwrightOperation";
import {
  Timeout,
  LocalDebugTaskLabel,
  LocalDebugError,
  LocalDebugTaskResult,
  Lang,
} from "../../utils/constants";
import { Env } from "../../utils/env";
import { validateFileExist } from "../../utils/commonUtils";
import { expect } from "chai";
import { LocalDebugTestContext } from "../localdebug/localdebugContext";
import { ModalDialog, VSBrowser } from "vscode-extension-tester";
import {
  RemoteDebugTestContext,
  provisionProject,
  deployProject,
  setSkuNameToStandard,
} from "../remotedebug/remotedebugContext";
import {
  execCommandIfExist,
  createNewProject,
} from "../../utils/vscodeOperation";
import { it } from "../../utils/it";

describe("Local Debug M365 Tests", function () {
  this.timeout(Timeout.localAndRemoteTestCase);
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
    "[auto] [Typescript] Local debug for SSO enabled personal tab project",
    {
      testPlanCaseId: 15277099,
      author: "kuojianlu@microsoft.com",
    },
    async () => {
      const localDebugTestContext = new LocalDebugTestContext("m365lp", {
        lang: Lang.TS,
      });
      await localDebugTestContext.before();
      const projectPath = path.resolve(
        localDebugTestContext.testRootFolder,
        localDebugTestContext.appName
      );
      validateFileExist(projectPath, "src/index.tsx");

      await startDebugging("Debug in Teams (Chrome)");

      try {
        await waitForTerminal(
          LocalDebugTaskLabel.StartBackend,
          LocalDebugTaskResult.FunctionStarted
        );
        await clearNotifications();
        await waitForTerminal(
          LocalDebugTaskLabel.StartFrontend,
          LocalDebugTaskResult.FrontendReady
        );
      } catch (error) {
        const errorMsg = error.toString();
        if (
          // skip can't find element
          errorMsg.includes(LocalDebugError.ElementNotInteractableError) ||
          // skip timeout
          errorMsg.includes(LocalDebugError.TimeoutError)
        ) {
          console.log("[skip error] ", error);
        } else {
          expect.fail(errorMsg);
        }
      }

      const teamsAppId = await localDebugTestContext.getTeamsAppId();
      const page = await initPage(
        localDebugTestContext.context!,
        teamsAppId,
        Env.username,
        Env.password,
        { projectPath: projectPath, env: "local" }
      );
      await localDebugTestContext.validateLocalStateForTab();
      await validateReactTab(page, Env.displayName, true);
      const m365AppId = await localDebugTestContext.getM365AppId();
      const url = `https://outlook.office.com/host/${m365AppId}/index?login_hint=${Env.username}`;
      await validateReactOutlookTab(page, url, Env.displayName, true);
      successFlagForLocal = true;
      await localDebugTestContext.after();
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
    }
  );

  it(
    "[auto] [Typescript] Remote debug for SSO enabled personal tab project",
    {
      testPlanCaseId: 15277096,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      const appNameCopySuffix = "copy";
      const remoteDebugTestContext = new RemoteDebugTestContext("tab");
      const testRootFolder = remoteDebugTestContext.testRootFolder;
      const appName = remoteDebugTestContext.appName;
      const newAppFolderName = appName + appNameCopySuffix;
      const projectPath = path.resolve(testRootFolder, newAppFolderName);
      await remoteDebugTestContext.before();
      //create tab project
      const driver = VSBrowser.instance.driver;
      try {
        await createNewProject("m365lp", appName, { lang: Lang.TS });
        await setSkuNameToStandard(projectPath);
        await driver.sleep(Timeout.shortTimeWait);
        await provisionProject(appName, projectPath);
        await deployProject(projectPath);
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
        await validateReactTab(page, Env.displayName, true);
        successFlagForRemote = true;
      } catch (error) {
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
          true,
          false,
          false
        );
        throw new Error("Error: " + error);
      }
    }
  );
});

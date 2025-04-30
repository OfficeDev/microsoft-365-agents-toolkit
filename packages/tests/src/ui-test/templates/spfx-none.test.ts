// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Anne Fu <v-annefu@microsoft.com>
 */
import { startDebugging, waitForTerminal } from "../../utils/vscodeOperation";
import {
  initPage,
  validateTeamsWorkbench,
  validateSpfx,
} from "../../utils/playwrightOperation";
import { LocalDebugTestContext } from "../localdebug/localdebugContext";
import { Env } from "../../utils/env";
import { it } from "../../utils/it";
import * as path from "path";
import { ModalDialog, VSBrowser } from "vscode-extension-tester";
import {
  CommandPaletteCommands,
  Timeout,
  Notification,
} from "../../utils/constants";
import {
  RemoteDebugTestContext,
  runDeploy,
} from "../remotedebug/remotedebugContext";
import {
  execCommandIfExist,
  getNotification,
  createNewProject,
  clearNotifications,
} from "../../utils/vscodeOperation";
import { cleanUpLocalProject } from "../../utils/cleanHelper";
import { validateFileExist } from "../../utils/commonUtils";

describe("SPFx local debug", function () {
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
    "[auto] Debug SPFx with none framework",
    {
      testPlanCaseId: 9454461,
      author: "v-annefu@microsoft.com",
    },
    async () => {
      const localDebugTestContext = new LocalDebugTestContext("spfx", {
        framework: "none",
      });
      await localDebugTestContext.before();
      const projectPath = path.resolve(
        localDebugTestContext.testRootFolder,
        localDebugTestContext.appName
      );
      await startDebugging("Teams workbench (Chrome)");

      // await waitForTerminal(LocalDebugTaskLabel.TabsNpmInstall);
      // await waitForTerminal("gulp trust-dev-cert");
      await waitForTerminal("gulp serve");

      const teamsAppId = await localDebugTestContext.getTeamsAppId();
      const page = await initPage(
        localDebugTestContext.context!,
        teamsAppId,
        Env.username,
        Env.password,
        { projectPath: projectPath, env: "local" }
      );
      await validateTeamsWorkbench(page, localDebugTestContext.appName);
      successFlagForLocal = true;
      await localDebugTestContext.after(false);
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
    "[auto] Create and run SPFx project with None framework",
    {
      testPlanCaseId: 9454331,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      const driver = VSBrowser.instance.driver;
      const remoteDebugTestContext = new RemoteDebugTestContext("spfx");
      const testRootFolder = remoteDebugTestContext.testRootFolder;
      const appName = remoteDebugTestContext.appName;
      const appNameCopySuffix = "copy";
      const newAppFolderName = appName + appNameCopySuffix;
      const projectPath = path.resolve(testRootFolder, newAppFolderName);
      await remoteDebugTestContext.before();
      await createNewProject("spfx", appName, { spfxFrameworkType: "None" });
      validateFileExist(projectPath, "src/src/index.ts");
      await clearNotifications();
      await execCommandIfExist(CommandPaletteCommands.ProvisionCommand);
      await driver.sleep(Timeout.spfxProvision);
      await getNotification(
        Notification.ProvisionSucceeded,
        Timeout.shortTimeWait
      );
      await runDeploy();

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

      // Validate app name is in the page
      await validateSpfx(page, { displayName: appName });
      successFlagForRemote = true;
      // Close the folder and cleanup local sample project
      await execCommandIfExist("Workspaces: Close Workspace", Timeout.webView);
      await cleanUpLocalProject(projectPath);
    }
  );
});

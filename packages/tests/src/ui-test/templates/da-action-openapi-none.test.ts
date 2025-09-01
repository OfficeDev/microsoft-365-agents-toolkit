// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Anne Fu <v-annefu@microsoft.com>
 */
import * as path from "path";
import { VSBrowser } from "vscode-extension-tester";
import {
  CommandPaletteCommands,
  Timeout,
  Notification,
} from "../../utils/constants";
import { RemoteDebugTestContext } from "../remotedebug/remotedebugContext";
import {
  execCommandIfExist,
  getNotification,
  createNewProject,
  clearNotifications,
} from "../../utils/vscodeOperation";
import { cleanUpLocalProject, cleanTeamsApp } from "../../utils/cleanHelper";
import { it } from "../../utils/it";
import {
  initCopilotPage,
  validatePrompt,
} from "../../utils/playwrightOperation";
import { Env } from "../../utils/env";

describe("Remote debug Tests", function () {
  this.timeout(Timeout.testAzureCase);
  let remoteDebugTestContext: RemoteDebugTestContext;
  let testRootFolder: string;
  let appName: string;
  const appNameCopySuffix = "copy";
  let newAppFolderName: string;
  let projectPath: string;

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    remoteDebugTestContext = new RemoteDebugTestContext("daOpenAPI");
    testRootFolder = remoteDebugTestContext.testRootFolder;
    appName = remoteDebugTestContext.appName;
    newAppFolderName = appName + appNameCopySuffix;
    projectPath = path.resolve(testRootFolder, newAppFolderName);
    await remoteDebugTestContext.before();
  });

  afterEach(async function () {
    this.timeout(Timeout.finishAzureTestCase);
    await remoteDebugTestContext.after();
    //Close the folder and cleanup local sample project
    await execCommandIfExist("Workspaces: Close Workspace", Timeout.webView);
    console.log(`[Successfully] start to clean up for ${projectPath}`);
    // uninstall Teams app
    cleanTeamsApp(appName), cleanUpLocalProject(projectPath);
  });

  it(
    "[auto][VSC]Create a basic declarative agent with API spec",
    {
      testPlanCaseId: 28941863,
      author: "v-annefu@microsoft.com",
    },
    async function () {
      const driver = VSBrowser.instance.driver;
      await createNewProject("daOpenAPI", appName, { apiAuthOption: "None" });
      await clearNotifications();
      await execCommandIfExist(CommandPaletteCommands.ProvisionCommand);
      await driver.sleep(Timeout.openAPIProvision);
      await getNotification(
        Notification.ProvisionSucceeded,
        Timeout.shortTimeWait
      );
      await clearNotifications();
      const teamsAppId = await remoteDebugTestContext.getTeamsAppId(
        projectPath
      );
      const page = await initCopilotPage(
        remoteDebugTestContext.context!,
        Env.username,
        Env.password
      );
      await driver.sleep(Timeout.longTimeWait);
      await validatePrompt(page, appName, {
        prompt: "List all repairs without auth",
        expected: "Oil",
        consent: false,
      });
    }
  );
});

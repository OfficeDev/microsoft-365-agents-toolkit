// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Helly Zhang <v-helzha@microsoft.com>
 */
import * as path from "path";
import { VSBrowser } from "vscode-extension-tester";
import { Lang, Timeout, ValidationContent } from "../../utils/constants";
import {
  RemoteDebugTestContext,
  setSkuNameToB1,
  provisionProject,
  deployProject,
} from "./remotedebugContext";
import {
  execCommandIfExist,
  createNewProject,
  clearNotifications,
} from "../../utils/vscodeOperation";
import { initPage, validateBasicTab } from "../../utils/playwrightOperation";
import { Env } from "../../utils/env";
import { it } from "../../utils/it";
import {
  cleanUpResourceGroup,
  createResourceGroup,
} from "../../utils/cleanHelper";

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
    remoteDebugTestContext = new RemoteDebugTestContext("tab");
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
    await remoteDebugTestContext.cleanUp(
      appName,
      projectPath,
      false,
      false,
      false
    );
  });

  it(
    "[auto] Delete resource group and re-provision for tab project",
    {
      testPlanCaseId: 10744678,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      //create tab project
      const driver = VSBrowser.instance.driver;
      await createNewProject("tabnsso", appName, { lang: Lang.TS });
      await setSkuNameToB1(projectPath);
      await driver.sleep(Timeout.shortTimeWait);
      await provisionProject(appName, projectPath);
      await clearNotifications();
      await cleanUpResourceGroup(appName, "dev");
      // wait for resource group to be deleted
      await driver.sleep(180 * 1000);
      await createResourceGroup(appName, "dev", "westus");
      // rerun provision
      await provisionProject(appName, projectPath, false);
      await deployProject(projectPath);
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
      await validateBasicTab(page, ValidationContent.Tab);
    }
  );
});

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Helly Zhang <v-helzha@microsoft.com>
 */
import * as path from "path";
import * as chai from "chai";
import { VSBrowser } from "vscode-extension-tester";
import { Timeout } from "../../utils/constants";
import {
  RemoteDebugTestContext,
  getAadObjectId,
  provisionProject,
} from "./remotedebugContext";
import {
  execCommandIfExist,
  createNewProject,
  runDeployAadAppManifest,
} from "../../utils/vscodeOperation";
import { Env } from "../../utils/env";
import { it } from "../../utils/it";
import { updateAadTemplate } from "../../utils/commonUtils";
import { GraphApiCleanHelper } from "../../utils/cleanHelper";

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
      true,
      false,
      false
    );
  });

  it(
    "[auto] AAD manifest feature VSCode E2E test - Deploy AAD App Manifest for remote debug",
    {
      testPlanCaseId: 16477901,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      //create tab project
      const driver = VSBrowser.instance.driver;
      await createNewProject("tab", appName);
      await provisionProject(appName, projectPath);

      await updateAadTemplate(projectPath, "-updated");
      await driver.sleep(Timeout.shortTimeWait);

      await runDeployAadAppManifest("dev");

      // await getNotification(
      //   "Your Azure Active Directory application has been successfully deployed."
      // );

      await driver.sleep(Timeout.longTimeWait);

      const cleanService = await GraphApiCleanHelper.create(
        Env.cleanTenantId,
        Env.cleanClientId,
        Env.username,
        Env.password
      );

      const aadObjectId = await getAadObjectId(projectPath);
      console.log(`get AAD ${aadObjectId}`);

      const aadInfo = await cleanService.getAad(aadObjectId);

      console.log(`get AAD ${aadInfo.displayName}`);
      const aadDisplayName = aadInfo.displayName as string;

      chai.expect(aadDisplayName).contains("updated");
    }
  );
});

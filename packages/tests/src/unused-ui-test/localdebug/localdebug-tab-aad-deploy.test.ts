// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Xiaofu Huang <xiaofu.huang@microsoft.com>
 */
import * as path from "path";
import * as chai from "chai";
import { VSBrowser } from "vscode-extension-tester";
import {
  clearNotifications,
  getNotification,
  runDeployAadAppManifest,
  startDebugging,
  stopDebugging,
  waitForTerminal,
} from "../../utils/vscodeOperation";
import { LocalDebugTestContext } from "./localdebugContext";
import {
  Timeout,
  LocalDebugTaskLabel,
  DebugItemSelect,
  LocalDebugTaskResult,
  LocalDebugError,
} from "../../utils/constants";
import { Env } from "../../utils/env";
import { it } from "../../utils/it";
import { validateFileExist, updateAadTemplate } from "../../utils/commonUtils";
import { GraphApiCleanHelper } from "../../utils/cleanHelper";

describe("Local Debug Tests", function () {
  this.timeout(Timeout.testCase);
  let localDebugTestContext: LocalDebugTestContext;

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    localDebugTestContext = new LocalDebugTestContext("tab");
    await localDebugTestContext.before();
  });

  afterEach(async function () {
    this.timeout(Timeout.finishTestCase);
    await localDebugTestContext.after();
  });

  it(
    "[auto] AAD manifest feature VSCode E2E test - Deploy AAD App Manifest for local debug",
    {
      testPlanCaseId: 16477393,
      author: "xiaofu.huang@microsoft.com",
    },
    async () => {
      const projectPath = path.resolve(
        localDebugTestContext.testRootFolder,
        localDebugTestContext.appName
      );
      validateFileExist(projectPath, "src/index.jsx");
      const driver = VSBrowser.instance.driver;

      await startDebugging(DebugItemSelect.DebugInTeamsUsingChrome);

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
          chai.expect.fail(errorMsg);
        }
      }

      await stopDebugging();

      await updateAadTemplate(projectPath, "-updated");
      await driver.sleep(Timeout.shortTimeWait);
      console.log("File aad.manifest.json is updated");

      await runDeployAadAppManifest("local");
      await driver.sleep(2 * Timeout.longTimeWait);
      console.log("AAD is updated");

      const cleanService = await GraphApiCleanHelper.create(
        Env.cleanTenantId,
        Env.cleanClientId,
        Env.username,
        Env.password
      );

      const aadObjectId = await localDebugTestContext.getAadObjectId();
      console.log(`get AAD ${aadObjectId}`);

      const aadInfo = await cleanService.getAad(aadObjectId);

      console.log(`get AAD ${aadInfo.displayName}`);
      const aadDisplayName = aadInfo.displayName as string;

      chai.expect(aadDisplayName).contains("updated");
    }
  );
});

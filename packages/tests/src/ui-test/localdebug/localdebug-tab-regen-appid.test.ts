// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Xiaofu Huang <xiaofu.huang@microsoft.com>
 */
import * as path from "path";
import {
  clearNotifications,
  startDebugging,
  stopDebugging,
  waitForTerminal,
} from "../../utils/vscodeOperation";
import { initPage, validateBasicTab } from "../../utils/playwrightOperation";
import { LocalDebugTestContext } from "./localdebugContext";
import {
  Timeout,
  LocalDebugTaskLabel,
  DebugItemSelect,
  ValidationContent,
  LocalDebugTaskResult,
} from "../../utils/constants";
import { Env } from "../../utils/env";
import { it } from "../../utils/it";
import { killPort, validateFileExist } from "../../utils/commonUtils";
import { cleanAppStudio } from "../../utils/cleanHelper";
import { ModalDialog, VSBrowser } from "vscode-extension-tester";

describe("Local Debug Tests", function () {
  this.timeout(Timeout.testCase);
  let localDebugTestContext: LocalDebugTestContext;

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    localDebugTestContext = new LocalDebugTestContext("tabnsso");
    await localDebugTestContext.before();
  });

  afterEach(async function () {
    this.timeout(Timeout.finishTestCase);
    await localDebugTestContext.after(false);
  });

  it(
    "[auto] Local debug: remove teams app from dev portal and re-run",
    {
      testPlanCaseId: 10822115,
      author: "xiaofu.huang@microsoft.com",
    },
    async () => {
      const projectPath = path.resolve(
        localDebugTestContext.testRootFolder,
        localDebugTestContext.appName
      );
      validateFileExist(projectPath, "src/app.js");
      const driver = VSBrowser.instance.driver;

      await startDebugging(DebugItemSelect.DebugInTeamsUsingChrome);

      await waitForTerminal(
        LocalDebugTaskLabel.StartApplication,
        LocalDebugTaskResult.FrontendStarted
      );

      await stopDebugging();
      await driver.sleep(Timeout.stopdebugging);
      try {
        await killPort(53000);
        console.log(`close port 53000 successfully`);
      } catch (error) {
        console.log(`close port 53000 failed`);
      }

      await cleanAppStudio(localDebugTestContext.appName);

      await driver.sleep(Timeout.shortTimeLoading);

      try {
        await clearNotifications();
        await startDebugging(DebugItemSelect.DebugInTeamsUsingChrome);
        await waitForTerminal(
          LocalDebugTaskLabel.StartApplication,
          LocalDebugTaskResult.FrontendStarted
        );
        // check if there is error "Could not attach to main target"
        await driver.sleep(Timeout.startdebugging);
        await waitForTerminal(
          LocalDebugTaskLabel.StartApplication,
          LocalDebugTaskResult.FrontendStarted
        );
      } catch {
        try {
          console.log(`Try to click "Cancel" button for error dialog`);
          const dialog = new ModalDialog();
          await dialog.pushButton("Cancel");
          await driver.sleep(Timeout.shortTimeLoading);
          console.log(
            `Clicked button "Cancel" for failing to attach to main target`
          );
          await stopDebugging();
          await driver.sleep(Timeout.stopdebugging);
        } catch {
          console.log(`No error for failing to attach to main target`);
        }

        try {
          await killPort(53000);
          console.log(`close port 53000 successfully`);
        } catch (error) {
          console.log(`close port 53000 failed`);
        }

        try {
          await startDebugging(DebugItemSelect.DebugInTeamsUsingChrome);
          await waitForTerminal(
            LocalDebugTaskLabel.StartApplication,
            LocalDebugTaskResult.FrontendStarted
          );
          // check if there is error "Debug Anyway"
          await driver.sleep(Timeout.startdebugging);
          await waitForTerminal(
            LocalDebugTaskLabel.StartApplication,
            LocalDebugTaskResult.FrontendStarted
          );
        } catch {
          console.log(`Try to click "Debug Anyway" button for error dialog`);
          const dialog = new ModalDialog();
          await dialog.pushButton("Debug Anyway");
          console.log(`Clicked button "Debug Anyway"`);
          await driver.sleep(Timeout.shortTimeLoading);
          await waitForTerminal(
            LocalDebugTaskLabel.StartApplication,
            LocalDebugTaskResult.FrontendStarted
          );
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
      await validateBasicTab(page, ValidationContent.Tab);
    }
  );
});

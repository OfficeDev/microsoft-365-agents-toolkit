// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Xiaofu Huang <xiaofu.huang@microsoft.com>
 */
import * as path from "path";
import {
  startDebugging,
  stopDebugging,
  waitForTerminal,
} from "../../utils/vscodeOperation";
import { initPage, validateEchoBot } from "../../utils/playwrightOperation";
import { LocalDebugTestContext } from "./localdebugContext";
import {
  Timeout,
  LocalDebugTaskLabel,
  LocalDebugTaskInfo,
  DebugItemSelect,
} from "../../utils/constants";
import { Env } from "../../utils/env";
import { it } from "../../utils/it";
import { killPort, validateFileExist } from "../../utils/commonUtils";
import { ModalDialog, VSBrowser } from "vscode-extension-tester";

describe("Local Debug Tests", function () {
  this.timeout(Timeout.testCase);
  let localDebugTestContext: LocalDebugTestContext;

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    localDebugTestContext = new LocalDebugTestContext("bot");
    await localDebugTestContext.before();
  });

  afterEach(async function () {
    this.timeout(Timeout.finishTestCase);
    await localDebugTestContext.after(false, true);
  });

  it(
    "[auto] Second press F5 to local debug for Bot successfully",
    {
      testPlanCaseId: 9795544,
      author: "xiaofu.huang@microsoft.com",
    },
    async function () {
      const projectPath = path.resolve(
        localDebugTestContext.testRootFolder,
        localDebugTestContext.appName
      );
      validateFileExist(projectPath, "index.js");
      const driver = VSBrowser.instance.driver;
      await startDebugging(DebugItemSelect.DebugInTeamsUsingChrome);

      await waitForTerminal(LocalDebugTaskLabel.StartLocalTunnel);
      try {
        await waitForTerminal(
          LocalDebugTaskLabel.StartBotApp,
          LocalDebugTaskInfo.AppListening
        );
        await stopDebugging();
        await driver.sleep(Timeout.stopdebugging);
        try {
          await killPort(3978);
          console.log(`close port 3978 successfully`);
        } catch (error) {
          console.log(`close port 3978 failed`);
        }
        await startDebugging(DebugItemSelect.DebugInTeamsUsingChrome);
        await waitForTerminal(LocalDebugTaskLabel.StartLocalTunnel);
        await waitForTerminal(
          LocalDebugTaskLabel.StartBotApp,
          LocalDebugTaskInfo.AppListening
        );

        // check if there is error "Could not attach to main target"
        await driver.sleep(Timeout.startdebugging);
        await waitForTerminal(
          LocalDebugTaskLabel.StartBotApp,
          LocalDebugTaskInfo.AppListening
        );
      } catch {
        const dialog = new ModalDialog();
        console.log(`click "Cancel" button for error dialog`);
        await dialog.pushButton("Cancel");
        await driver.sleep(Timeout.shortTimeLoading);
        console.log(
          `Clicked button "Cancel" for failing to attach to main target`
        );
        await stopDebugging();
        await driver.sleep(Timeout.stopdebugging);
        try {
          await killPort(3978);
          console.log(`close port 3978 successfully`);
        } catch (error) {
          console.log(`close port 3978 failed`);
        }
        await startDebugging(DebugItemSelect.DebugInTeamsUsingChrome);
        try {
          await waitForTerminal(
            LocalDebugTaskLabel.StartBotApp,
            LocalDebugTaskInfo.AppListening
          );

          // check if there is error "Debug Anyway"
          await driver.sleep(Timeout.startdebugging);
          await waitForTerminal(
            LocalDebugTaskLabel.StartBotApp,
            LocalDebugTaskInfo.AppListening
          );
        } catch {
          const dialog = new ModalDialog();
          console.log(`click "Debug Anyway" button for error dialog`);
          await dialog.pushButton("Debug Anyway");
          console.log(`Clicked button "Debug Anyway"`);
          await driver.sleep(Timeout.shortTimeLoading);
          await waitForTerminal(
            LocalDebugTaskLabel.StartBotApp,
            LocalDebugTaskInfo.AppListening
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
      await localDebugTestContext.validateLocalStateForBot();
      await validateEchoBot(page);
    }
  );
});

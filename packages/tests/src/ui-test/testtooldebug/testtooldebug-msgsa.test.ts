// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Helly Zhang <v-helzha@microsoft.com>
 */

import * as path from "path";
import os from "os";
import { startDebugging, waitForTerminal } from "../../utils/vscodeOperation";
import { LocalDebugTestContext } from "../localdebug/localdebugContext";
import {
  Timeout,
  LocalDebugTaskLabel,
  LocalDebugTaskInfo,
  DebugItemSelect,
  LocalDebugTaskLabel2,
} from "../../utils/constants";
import { it } from "../../utils/it";
import { validateFileExist } from "../../utils/commonUtils";
import { VSBrowser } from "vscode-extension-tester";
import { expect } from "chai";
import { getScreenshotName } from "../../utils/nameUtil";
import { validateNpm } from "../../utils/testToolValidations";

describe("Test Tool Debug Tests", function () {
  this.timeout(Timeout.testAzureCase);
  let localDebugTestContext: LocalDebugTestContext;
  let successFlag = true;
  let errorMessage = "";

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    localDebugTestContext = new LocalDebugTestContext("msgsa");
    await localDebugTestContext.before();
  });

  after(async function () {
    this.timeout(Timeout.finishTestCase);
    await localDebugTestContext.after(false, false);
    setTimeout(() => {
      if (os.type() === "Windows_NT") {
        if (successFlag) process.exit(0);
        else process.exit(1);
      }
    }, 30000);
  });

  it(
    "[ME] Debug Message Extension Search Command in Test Tool",
    {
      testPlanCaseId: 27548668,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      try {
        const projectPath = path.resolve(
          localDebugTestContext.testRootFolder,
          localDebugTestContext.appName
        );
        validateFileExist(projectPath, "src/index.js");
        const driver = VSBrowser.instance.driver;

        // local debug in Test Tool
        await startDebugging(DebugItemSelect.DebugInTestTool);

        await waitForTerminal(
          LocalDebugTaskLabel.StartBotApp,
          LocalDebugTaskInfo.StartBotInfo
        );

        await waitForTerminal(LocalDebugTaskLabel2.StartTestTool);

        await driver.sleep(Timeout.startdebugging);

        await validateNpm(localDebugTestContext.context!, {
          npmName: "axios",
          appName: localDebugTestContext.appName,
        });
      } catch (error) {
        successFlag = false;
        errorMessage = "[Error]: " + error;
        await VSBrowser.instance.takeScreenshot(getScreenshotName("error"));
        await VSBrowser.instance.driver.sleep(Timeout.playwrightDefaultTimeout);
      }
      expect(successFlag, errorMessage).to.true;
      console.log("debug finish!");
    }
  );
});

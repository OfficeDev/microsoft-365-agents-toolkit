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
import { validateWelcomeAndReplyBot } from "../../utils/testToolValidations";

// TODO: Change preview test to normal test before rc release
describe("Command And Response Bot Local Debug Tests", function () {
  this.timeout(Timeout.testAzureCase);
  let localDebugTestContext: LocalDebugTestContext;
  let successFlag = true;
  let errorMessage = "";

  const oldEnv = Object.assign({}, process.env);
  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    localDebugTestContext = new LocalDebugTestContext("crbot");
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
    "[auto] Local debug using Test Tool for Command and Response Bot App",
    {
      testPlanCaseId: 25666171,
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

        await validateWelcomeAndReplyBot(localDebugTestContext.context!, {
          hasWelcomeMessage: false,
          hasCommandReplyValidation: true,
          botCommand: "helloWorld",
          expectedWelcomeMessage: "Welcome to the Command Bot!",
          expectedReplyMessage: "Your Hello World App is Running",
          timeout: Timeout.longTimeWait,
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

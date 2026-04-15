// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */
import * as path from "path";
import { startDebugging, waitForTerminal } from "../../utils/vscodeOperation";
import {
  initPlaygroundPage,
  validateEchoBotInPlayground,
} from "../../utils/playwrightOperation";
import { PlaygroundTestContext } from "./playgroundContext";
import {
  Timeout,
  LocalDebugTaskLabel,
  LocalDebugTaskInfo,
  DebugItemSelect,
} from "../../utils/constants";
import { it } from "../../utils/it";
import { validateFileExist } from "../../utils/commonUtils";
import { VSBrowser } from "vscode-extension-tester";

describe("Local Debug Tests", function () {
  this.timeout(Timeout.testCase);
  let playgroundTestContext: PlaygroundTestContext;

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    playgroundTestContext = new PlaygroundTestContext("bot");
    await playgroundTestContext.before();
  });

  afterEach(async function () {
    this.timeout(Timeout.finishTestCase);
    await playgroundTestContext.after();
  });

  it(
    "[auto] Second press F5 to playground debug for Bot successfully",
    {
      testPlanCaseId: 36232743,
      author: "v-ivanchen@microsoft.com",
    },
    async function () {
      const projectPath = path.resolve(
        playgroundTestContext.testRootFolder,
        playgroundTestContext.appName,
      );
      validateFileExist(projectPath, "index.js");
      const driver = VSBrowser.instance.driver;

      await startDebugging(DebugItemSelect.DebugInAgentsPlayground);
      await waitForTerminal(
        LocalDebugTaskLabel.StartApplicationPlayground,
        LocalDebugTaskInfo.AppListening,
      );

      // check if there is error "Could not attach to main target"
      await driver.sleep(Timeout.startdebugging);

      await waitForTerminal(
        LocalDebugTaskLabel.StartAgentsPlayground,
        LocalDebugTaskInfo.PlaygroundStart,
      );

      const page = await initPlaygroundPage(playgroundTestContext.context!);
      await validateEchoBotInPlayground(page, { botCommand: "Hi" });
    },
  );
});

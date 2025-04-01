// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Helly Zhang <v-helzha@microsoft.com>
 */
import { startDebugging, waitForTerminal } from "../../utils/vscodeOperation";
import {
  initPage,
  switchToTab,
  validateSpfx,
  validateTeamsWorkbench,
} from "../../utils/playwrightOperation";
import { LocalDebugTestContext } from "./localdebugContext";
import { Timeout } from "../../utils/constants";
import { Env } from "../../utils/env";
import { it } from "../../utils/it";
import {
  configSpfxGlobalEnv,
  generateYoSpfxProject,
} from "../../utils/commonUtils";
import * as path from "path";

describe("SPFx local debug", function () {
  this.timeout(Timeout.testCase);
  let localDebugTestContext: LocalDebugTestContext;

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    await configSpfxGlobalEnv();
    await generateYoSpfxProject({
      solutionName: "existingspfx",
      componentName: "helloworld1",
    });
    await generateYoSpfxProject({
      existingSolutionName: "existingspfx",
      componentName: "helloworld2",
    });
    localDebugTestContext = new LocalDebugTestContext("spfximport", {
      existingSpfxFolder: "existingspfx",
    });
    await localDebugTestContext.before();
  });

  afterEach(async function () {
    this.timeout(Timeout.finishTestCase);
    await localDebugTestContext.after(false);
  });

  it(
    "[auto] Import existing SPFx solution with multiple web parts",
    {
      testPlanCaseId: 24434596,
      author: "v-helzha@microsoft.com",
    },
    async () => {
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
      await validateTeamsWorkbench(page, "helloworld1");
      await switchToTab(page, "helloworld2");
      await validateSpfx(page, {
        displayName: "helloworld2",
      });
    }
  );
});

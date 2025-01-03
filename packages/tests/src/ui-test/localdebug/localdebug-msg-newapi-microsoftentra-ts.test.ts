// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Anne Fu <v-annefu@microsoft.com>
 */
import * as path from "path";
import { startDebugging, waitForTerminal } from "../../utils/vscodeOperation";
import {
  initNoAddappPage,
  validateApiMeResult,
} from "../../utils/playwrightOperation";
import { LocalDebugTestContext } from "./localdebugContext";
import { Timeout, LocalDebugTaskLabel } from "../../utils/constants";
import { Env } from "../../utils/env";
import { it } from "../../utils/it";
import { validateFileExist } from "../../utils/commonUtils";

describe("Local Debug Tests", function () {
  this.timeout(Timeout.testCase);
  let localDebugTestContext: LocalDebugTestContext;

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    localDebugTestContext = new LocalDebugTestContext("msgmicroentra", {
      lang: "typescript",
    });
    await localDebugTestContext.before();
  });

  afterEach(async function () {
    this.timeout(Timeout.finishTestCase);
    await localDebugTestContext.after(false, true);
  });

  it(
    "[Typescript] Local debug for API Message Extension with Microsoft Entra auth",
    {
      testPlanCaseId: 28665881,
      author: "v-annefu@microsoft.com",
    },
    async function () {
      const projectPath = path.resolve(
        localDebugTestContext.testRootFolder,
        localDebugTestContext.appName
      );
      validateFileExist(projectPath, "src/functions/repair.ts");
      await startDebugging("Debug in Teams (Chrome)");
      await waitForTerminal(LocalDebugTaskLabel.StartLocalTunnel);
      await waitForTerminal(
        LocalDebugTaskLabel.StartBackend,
        "Worker process started and initialized"
      );
      const teamsAppId = await localDebugTestContext.getTeamsAppId();
      const page = await initNoAddappPage(
        localDebugTestContext.context!,
        teamsAppId,
        Env.username,
        Env.password
      );
      await validateApiMeResult(page, localDebugTestContext.appName);
    }
  );
});

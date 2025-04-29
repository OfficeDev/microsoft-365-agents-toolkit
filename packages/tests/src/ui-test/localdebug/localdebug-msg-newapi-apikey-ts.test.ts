// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Anne Fu <v-annefu@microsoft.com>
 */
import * as path from "path";
import * as fs from "fs-extra";
import { startDebugging, waitForTerminal } from "../../utils/vscodeOperation";
import {
  initNoAddappPage,
  initPage,
  validateApiMeResult,
} from "../../utils/playwrightOperation";
import { LocalDebugTestContext } from "./localdebugContext";
import { Timeout, LocalDebugTaskLabel, Lang } from "../../utils/constants";
import { Env } from "../../utils/env";
import { it } from "../../utils/it";
import { validateFileExist } from "../../utils/commonUtils";

describe("Local Debug Tests", function () {
  this.timeout(Timeout.testCase);
  let localDebugTestContext: LocalDebugTestContext;

  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    localDebugTestContext = new LocalDebugTestContext("msgapikey", {
      lang: Lang.TS,
    });
    await localDebugTestContext.before();
  });

  afterEach(async function () {
    this.timeout(Timeout.finishTestCase);
    await localDebugTestContext.after(false, true);
  });

  it(
    "[Typescript] Local debug for API Message Extension with API key auth",
    {
      testPlanCaseId: 28289198,
      author: "v-annefu@microsoft.com",
    },
    async function () {
      const projectPath = path.resolve(
        localDebugTestContext.testRootFolder,
        localDebugTestContext.appName
      );
      validateFileExist(projectPath, "src/functions/repair.ts");
      const userFile = path.resolve(projectPath, "env", ".env.local.user");
      const SECRET_API_KEY = "SECRET_API_KEY=gbxEWvk4p3sg";
      const KEY = "\n" + SECRET_API_KEY;
      fs.appendFileSync(userFile, KEY);
      console.log("add SECRET_API_KEY=yourapikey to .env file");
      await startDebugging("Debug in Teams (Chrome)");
      await waitForTerminal(LocalDebugTaskLabel.StartLocalTunnel);
      await waitForTerminal(
        LocalDebugTaskLabel.StartBackend,
        "Worker process started and initialized"
      );
      const teamsAppId = await localDebugTestContext.getTeamsAppId();
      const page = await initPage(
        localDebugTestContext.context!,
        teamsAppId,
        Env.username,
        Env.password,
        { projectPath: projectPath, env: "local" }
      );
      await validateApiMeResult(page, localDebugTestContext.appName);
    }
  );
});

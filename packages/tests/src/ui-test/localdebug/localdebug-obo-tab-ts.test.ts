// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Kuojian Lu <kuojianlu@microsoft.com>
 */
import * as path from "path";
import {
  clearNotifications,
  startDebugging,
  waitForTerminal,
} from "../../utils/vscodeOperation";
import {
  initPage,
  validateReactOutlookTab,
  validateReactTab,
} from "../../utils/playwrightOperation";
import { LocalDebugTestContext } from "./localdebugContext";
import {
  Timeout,
  LocalDebugTaskLabel,
  LocalDebugError,
  LocalDebugTaskResult,
} from "../../utils/constants";
import { Env } from "../../utils/env";
import { it } from "../../utils/it";
import { validateFileExist } from "../../utils/commonUtils";
import { expect } from "chai";

describe("Local Debug M365 Tests", function () {
  this.timeout(Timeout.testCase);
  let localDebugTestContext: LocalDebugTestContext;

  beforeEach(async function () {
    process.env.TEAMSFX_M365_APP = "true";
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    localDebugTestContext = new LocalDebugTestContext("m365lp", {
      lang: "typescript",
    });
    await localDebugTestContext.before();
  });

  afterEach(async function () {
    process.env.TEAMSFX_M365_APP = "false";
    this.timeout(Timeout.finishTestCase);
    await localDebugTestContext.after();
  });

  it(
    "[auto] [Typescript] Local debug for SSO enabled personal tab project",
    {
      testPlanCaseId: 15277099,
      author: "kuojianlu@microsoft.com",
    },
    async () => {
      const projectPath = path.resolve(
        localDebugTestContext.testRootFolder,
        localDebugTestContext.appName
      );
      validateFileExist(projectPath, "src/index.tsx");

      await startDebugging("Debug in Teams (Chrome)");

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
          expect.fail(errorMsg);
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
      await localDebugTestContext.validateLocalStateForTab();
      await validateReactTab(page, Env.displayName, true);
      const m365AppId = await localDebugTestContext.getM365AppId();
      const url = `https://outlook.office.com/host/${m365AppId}/index?login_hint=${Env.username}`;
      await validateReactOutlookTab(page, url, Env.displayName, true);
    }
  );
});

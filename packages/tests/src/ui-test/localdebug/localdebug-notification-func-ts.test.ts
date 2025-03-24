// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * @author Aocheng Wang <aochengwang@microsoft.com>
 */
import * as path from "path";
import {
  startDebugging,
  stopDebugging,
  waitForTerminal,
} from "../../utils/vscodeOperation";
import {
  initPage,
  validateNotificationBot,
} from "../../utils/playwrightOperation";
import { LocalDebugTestContext } from "./localdebugContext";
import {
  Timeout,
  LocalDebugTaskLabel,
  DebugItemSelect,
} from "../../utils/constants";
import { Env } from "../../utils/env";
import { it } from "../../utils/it";
import { validateFileExist } from "../../utils/commonUtils";
import { ModalDialog, VSBrowser } from "vscode-extension-tester";

// TODO: Change preview test to normal test before rc release
describe("Func Hosted Notification Bot Local Debug Tests", function () {
  this.timeout(Timeout.testCase);
  let localDebugTestContext: LocalDebugTestContext;

  const oldEnv = Object.assign({}, process.env);
  beforeEach(async function () {
    // ensure workbench is ready
    this.timeout(Timeout.prepareTestCase);
    localDebugTestContext = new LocalDebugTestContext("funcNoti", {
      lang: "typescript",
    });
    await localDebugTestContext.before();
  });

  afterEach(async function () {
    process.env = oldEnv;
    this.timeout(Timeout.finishTestCase);
    await localDebugTestContext.after(false, true);
  });

  it(
    "[auto] [Typescript] Local debug Func Hosted Notification Bot App",
    {
      testPlanCaseId: 15277351,
      author: "aochengwang@microsoft.com",
    },
    async function () {
      const projectPath = path.resolve(
        localDebugTestContext.testRootFolder,
        localDebugTestContext.appName
      );
      validateFileExist(projectPath, "src/httpTrigger.ts");
      const driver = VSBrowser.instance.driver;
      await startDebugging(DebugItemSelect.DebugInTeamsUsingChrome);
      await waitForTerminal(LocalDebugTaskLabel.StartLocalTunnel);
      try {
        await waitForTerminal(
          "Start Azurite emulator",
          "Azurite Blob service is successfully listening"
        );
        await waitForTerminal(
          LocalDebugTaskLabel.StartBotApp,
          "Worker process started and initialized"
        );
      } catch {
        const dialog = new ModalDialog();
        console.log("click Cancel button for error dialog");
        await dialog.pushButton("Cancel");
        await driver.sleep(Timeout.shortTimeLoading);
        console.log(
          "Clicked button Cancel for failing to attach to main target"
        );
        await stopDebugging();
        await startDebugging(DebugItemSelect.DebugInTeamsUsingChrome);
        try {
          await waitForTerminal(
            LocalDebugTaskLabel.StartBotApp,
            "Worker process started and initialized"
          );
        } catch {
          const dialog = new ModalDialog();
          console.log("click Cancel button for error dialog");
          await dialog.pushButton("Debug Anyway");
          console.log("Clicked button Debug Anyway");
          await driver.sleep(Timeout.shortTimeLoading);
          await waitForTerminal(
            LocalDebugTaskLabel.StartBotApp,
            "Worker process started and initialized"
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
      await validateNotificationBot(
        page,
        "http://127.0.0.1:3978/api/notification"
      );
    }
  );
});

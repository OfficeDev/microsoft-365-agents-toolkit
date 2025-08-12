// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import path from "path";
import {
  LocalDebugTestContext,
  LocalDebugTestName,
} from "../localdebug/localdebugContext";
import { initDebugPort, validateFileExist } from "../../utils/commonUtils";
import {
  createNewProject,
  execCommandIfExist,
  startDebugging,
  waitForTerminal,
} from "../../utils/vscodeOperation";
import {
  AppType,
  DebugItemSelect,
  LocalDebugTaskInfo,
  LocalDebugTaskLabel,
  Timeout,
  Lang,
} from "../../utils/constants";
import { expect } from "chai";
import { initPage } from "../../utils/playwrightOperation";
import { Env } from "../../utils/env";
import { Executor } from "../../utils/executor";
import { ModalDialog, VSBrowser } from "vscode-extension-tester";
import { getScreenshotName } from "../../utils/nameUtil";
import { ChildProcess, ChildProcessWithoutNullStreams } from "child_process";
import {
  deployProject,
  provisionProject,
  RemoteDebugTestContext,
} from "../remotedebug/remotedebugContext";

export async function msgHappyPathTestForLocalDebug(
  capability: LocalDebugTestName,
  options: {
    lang: Lang;
    successFlag: {
      successFlagForLocal: boolean;
      successFlagForRemote: boolean;
    };
    fileValidation?: string;
    localDebugTaskLabel?: string;
    localDebugTaskInfo?: string;
    validationFn: (page: any, options: { appName: string }) => Promise<void>;
  }
) {
  let errorMessage = "";
  let devtunnelProcess!: ChildProcessWithoutNullStreams | null;
  let debugProcess!: ChildProcess | null;
  const localDebugTestContext = new LocalDebugTestContext(capability, {
    lang: options.lang,
  });
  await localDebugTestContext.before();
  try {
    const projectPath = path.resolve(
      localDebugTestContext.testRootFolder,
      localDebugTestContext.appName
    );
    if (options.lang === "JavaScript") {
      validateFileExist(projectPath, options.fileValidation || "src/index.js");
    } else if (options.lang === "TypeScript") {
      validateFileExist(projectPath, options.fileValidation || "src/index.ts");
    }
    // local debug
    console.log("======= debug with ttk ========");
    await startDebugging(DebugItemSelect.DebugInTeamsUsingChrome);
    await waitForTerminal(LocalDebugTaskLabel.StartLocalTunnel);
    await waitForTerminal(
      options?.localDebugTaskLabel || LocalDebugTaskLabel.StartBotApp,
      options?.localDebugTaskInfo || LocalDebugTaskInfo.AppListening
    );

    const teamsAppId = await localDebugTestContext.getTeamsAppId();
    expect(teamsAppId).to.not.be.empty;
    {
      const page = await initPage(
        localDebugTestContext.context!,
        teamsAppId,
        Env.username,
        Env.password,
        {
          projectPath: projectPath,
          teamsAppName: localDebugTestContext.appName,
          env: "local",
          searchApp: false,
        }
      );
      await localDebugTestContext.validateLocalStateForBot();
      await options.validationFn(page, {
        appName: localDebugTestContext.appName,
      });
    }
    // cli preview
    console.log("======= debug with cli ========");
    let tunnelName = "";
    let devtunnelProcess: ChildProcessWithoutNullStreams | null = null;
    let debugProcess: ChildProcess | null = null;
    const tunnel = Executor.debugBotFunctionPreparation(projectPath);
    tunnelName = tunnel.tunnelName;
    devtunnelProcess = tunnel.devtunnelProcess;
    await new Promise((resolve) => setTimeout(resolve, 60 * 1000));
    {
      const { success } = await Executor.provision(
        projectPath,
        "local",
        true,
        "DeprecationWarning"
      );
      expect(success).to.be.true;
      console.log(`[Successfully] provision for ${projectPath}`);
    }
    {
      const { success } = await Executor.deploy(projectPath, "local");
      expect(success).to.be.true;
      console.log(`[Successfully] deploy for ${projectPath}`);
    }

    debugProcess = Executor.debugProject(
      projectPath,
      "local",
      true,
      process.env,
      (data) => {
        if (data) {
          console.log(data);
        }
      },
      (error) => {
        const errorMsg = error.toString();
      }
    );
    {
      const page = await initPage(
        localDebugTestContext.context!,
        teamsAppId,
        Env.username,
        Env.password,
        {
          projectPath: projectPath,
          teamsAppName: localDebugTestContext.appName,
          env: "local",
          searchApp: false,
          loggedIn: true, // to avoid re-login
        }
      );
      await localDebugTestContext.validateLocalStateForBot();
      await options.validationFn(page, {
        appName: localDebugTestContext.appName,
      });
      options.successFlag.successFlagForLocal = true;
      console.log(
        "successFlagForLocal: ",
        options.successFlag.successFlagForLocal
      );
    }
  } catch (error) {
    errorMessage = "[Error]: " + error;
    console.log(errorMessage);
    await VSBrowser.instance.takeScreenshot(getScreenshotName("error"));
    await VSBrowser.instance.driver.sleep(Timeout.playwrightDefaultTimeout);
  }
  // kill process
  await Executor.closeProcess(debugProcess);
  await Executor.closeProcess(devtunnelProcess);
  await initDebugPort();
  console.log("debug finish!");

  await localDebugTestContext.after(false, true);
  try {
    //Close the folder and cleanup local sample project
    await execCommandIfExist("Workspaces: Close Workspace", Timeout.webView);
  } catch {
    const dialog = new ModalDialog();
    console.log(`Click "Cancel" button if it exists`);
    await dialog.pushButton("Cancel");
    console.log(`Clicked button "Cancel"`);
    await execCommandIfExist("Workspaces: Close Workspace", Timeout.webView);
  }
  expect(options.successFlag.successFlagForLocal, errorMessage).to.true;
}

export async function msgHappyPathTestForRemoteDebug(
  capability: AppType,
  options: {
    lang: Lang;
    successFlag: {
      successFlagForLocal: boolean;
      successFlagForRemote: boolean;
    };
    fileValidation?: string;
    validationFn: (page: any, options: { appName: string }) => Promise<void>;
  }
) {
  const remoteDebugTestContext = new RemoteDebugTestContext("msg");
  const testRootFolder = remoteDebugTestContext.testRootFolder;
  const appName = remoteDebugTestContext.appName;
  const appNameCopySuffix = "copy";
  const newAppFolderName = appName + appNameCopySuffix;
  const projectPath = path.resolve(testRootFolder, newAppFolderName);
  await remoteDebugTestContext.before();
  const driver = VSBrowser.instance.driver;
  await createNewProject(capability, appName, {
    lang: options.lang,
  });
  if (options.lang === Lang.JS) {
    validateFileExist(projectPath, options.fileValidation || "src/index.js");
  } else if (options.lang === Lang.TS) {
    validateFileExist(projectPath, options.fileValidation || "src/index.ts");
  }
  await provisionProject(appName, projectPath);
  try {
    await deployProject(projectPath, Timeout.botDeploy);
    const teamsAppId = await remoteDebugTestContext.getTeamsAppId(projectPath);
    const page = await initPage(
      remoteDebugTestContext.context!,
      teamsAppId,
      Env.username,
      Env.password,
      {
        projectPath: projectPath,
        env: "dev",
        teamsAppName: appName,
        searchApp: false,
      }
    );
    await driver.sleep(Timeout.longTimeWait);
    await options.validationFn(page, {
      appName: appName,
    });
    options.successFlag.successFlagForRemote = true;
    console.log(
      "successFlagForRemote: ",
      options.successFlag.successFlagForRemote
    );
  } catch (error) {
    //Close the folder and cleanup local sample project
    await execCommandIfExist("Workspaces: Close Workspace", Timeout.webView);
    console.log(`[Successfully] start to clean up for ${projectPath}`);
    await remoteDebugTestContext.cleanUp(
      appName,
      projectPath,
      false,
      true,
      false
    );
    throw new Error("[Error]: " + error);
  }
}

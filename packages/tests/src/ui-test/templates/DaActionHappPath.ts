// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as path from "path";
import { ModalDialog, VSBrowser } from "vscode-extension-tester";
import {
  Lang,
  Timeout,
  AppType,
  DebugItemSelect,
  LocalDebugTaskLabel,
} from "../../utils/constants";
import {
  RemoteDebugTestContext,
  provisionProject,
  deployProject,
} from "../remotedebug/remotedebugContext";
import {
  execCommandIfExist,
  createNewProject,
  startDebugging,
  waitForTerminal,
} from "../../utils/vscodeOperation";
import { initCopilotPage } from "../../utils/playwrightOperation";
import { Env } from "../../utils/env";
import { editDotEnvFile, validateFileExist } from "../../utils/commonUtils";
import {
  LocalDebugTestContext,
  LocalDebugTestName,
} from "../localdebug/localdebugContext";
import { getScreenshotName } from "../../utils/nameUtil";
import { expect } from "chai";

export async function daActionHappPathTestForLocalDebug(
  capability: LocalDebugTestName,
  options: {
    lang: Lang;
    apiAuth: "none" | "api-key" | "microsoft-entra" | "oauth";
    successFlag: {
      successFlagForLocal: boolean;
      successFlagForRemote: boolean;
    };
    fileValidation?: string;
    validationFn: (
      page: any,
      options: {
        appName: string;
        prompt?: string;
        expected?: string;
        consent?: boolean;
      }
    ) => Promise<void>;
    prompt?: string;
    expected?: string;
    consent?: boolean;
  }
) {
  let errorMessage = "";
  const localDebugTestContext = new LocalDebugTestContext(capability, {
    lang: options.lang,
    apiAuth: options.apiAuth,
  });
  await localDebugTestContext.before();
  const driver = VSBrowser.instance.driver;
  try {
    const projectPath = path.resolve(
      localDebugTestContext.testRootFolder,
      localDebugTestContext.appName
    );
    if (options.lang === Lang.JS) {
      validateFileExist(
        projectPath,
        options.fileValidation || "src/functions/repairs.js"
      );
    } else if (options.lang === Lang.TS) {
      validateFileExist(
        projectPath,
        options.fileValidation || "src/functions/repairs.ts"
      );
    }
    if (options.apiAuth === "api-key") {
      const envPath = path.resolve(projectPath, "env/.env.local.user");
      editDotEnvFile(envPath, "SECRET_API_KEY", "fakekey-Mp7t");
      console.log("API Key set in .env.local.user");
    }
    // local debug
    console.log("======= debug with ttk ========");
    await startDebugging(DebugItemSelect.DebugInCopilotUsingChrome);
    await waitForTerminal(LocalDebugTaskLabel.StartLocalTunnel);
    await waitForTerminal(
      LocalDebugTaskLabel.StartBackend,
      "Worker process started and initialized"
    );
    await driver.sleep(Timeout.longTimeWait);
    const teamsAppId = await localDebugTestContext.getTeamsAppId();
    expect(teamsAppId).to.not.be.empty;
    {
      await driver.sleep(Timeout.longTimeWait);
      const copilotAgentName = `${localDebugTestContext.appName}local`;
      const page = await initCopilotPage(
        localDebugTestContext.context!,
        Env.username,
        Env.password,
        { copilotAgentName: copilotAgentName }
      );
      await options.validationFn(page, {
        appName: copilotAgentName,
        prompt: options?.prompt,
        expected: options?.expected,
        consent: options?.consent,
      });
    }
    options.successFlag.successFlagForLocal = true;
    console.log(
      "successFlagForLocal: ",
      options.successFlag.successFlagForLocal
    );
  } catch (error) {
    errorMessage = "[Error]: " + error;
    console.log(errorMessage);
    await VSBrowser.instance.takeScreenshot(getScreenshotName("error"));
    await VSBrowser.instance.driver.sleep(Timeout.playwrightDefaultTimeout);
  }

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

export async function daActionHappPathTestForRemoteDebug(
  capability: AppType,
  options: {
    lang: Lang;
    authOption: "None" | "API Key" | "Microsoft Entra" | "OAuth";
    successFlag: {
      successFlagForLocal: boolean;
      successFlagForRemote: boolean;
    };
    fileValidation?: string;
    validationFn: (
      page: any,
      options: {
        appName: string;
        prompt?: string;
        expected?: string;
        consent?: boolean;
      }
    ) => Promise<void>;
    prompt?: string;
    expected?: string;
    consent?: boolean;
  }
) {
  const remoteDebugTestContext = new RemoteDebugTestContext("danone");
  const testRootFolder = remoteDebugTestContext.testRootFolder;
  const appName = remoteDebugTestContext.appName;
  const appNameCopySuffix = "copy";
  const newAppFolderName = appName + appNameCopySuffix;
  const projectPath = path.resolve(testRootFolder, newAppFolderName);
  await remoteDebugTestContext.before();
  const driver = VSBrowser.instance.driver;

  await createNewProject(capability, appName, {
    lang: options.lang,
    authOption: options.authOption || "None",
  });
  if (options.authOption === "API Key") {
    const envPath = path.resolve(projectPath, "env/.env.dev.user");
    editDotEnvFile(envPath, "SECRET_API_KEY", "fakekey-Mp7t");
    console.log("API Key set in .env.dev.user");
  }
  if (options.lang === Lang.JS) {
    validateFileExist(
      projectPath,
      options.fileValidation || "src/functions/repairs.js"
    );
  } else if (options.lang === Lang.TS) {
    validateFileExist(
      projectPath,
      options.fileValidation || "src/functions/repairs.ts"
    );
  }
  await provisionProject(appName, projectPath);
  try {
    await deployProject(projectPath, Timeout.botDeploy);
    const teamsAppId = await remoteDebugTestContext.getTeamsAppId(projectPath);
    await driver.sleep(Timeout.longTimeWait);
    const copilotAgentName = `${appName}dev`;
    const page = await initCopilotPage(
      remoteDebugTestContext.context!,
      Env.username,
      Env.password,
      { copilotAgentName: copilotAgentName }
    );
    await options.validationFn(page, {
      appName: copilotAgentName,
      prompt: options?.prompt,
      expected: options?.expected,
      consent: options?.consent,
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
      false,
      false
    );
    throw new Error("[Error]: " + error);
  }
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as path from "path";
import * as fs from "fs-extra";
import { VSBrowser } from "vscode-extension-tester";
import { Browser, BrowserContext, chromium } from "playwright";
import { Timeout, TestFilePath } from "../utils/constants";
import { ensureExtensionActivated } from "../utils/vscodeOperation";
import {
  cleanAppStudio,
  cleanTeamsApp,
  DevTunnelCleanHelper,
  GraphApiCleanHelper,
  cleanUpResourceGroup,
  deleteResourceGroupByName,
} from "../utils/cleanHelper";
import { getAppName, getScreenshotName } from "../utils/nameUtil";
import { dotenvUtil } from "../utils/envUtil";
import { Env } from "../utils/env";

export class TestContext {
  public browser?: Browser;
  public context?: BrowserContext;
  public testRootFolder: string;
  public appName: string;

  constructor(testName: string) {
    this.testRootFolder = path.resolve(__dirname, "../../resource/");
    this.appName = getAppName(testName.split("-")[0]);
  }

  public async before(): Promise<void> {
    await fs.ensureDir(this.testRootFolder);
    await VSBrowser.instance.waitForWorkbench();
    this.browser = await chromium.launch({
      headless: false,
      slowMo: 100,
      timeout: Timeout.chromiumLaunchTimeout,
    });
    this.context = await this.browser.newContext({ ignoreHTTPSErrors: true });
    await VSBrowser.instance.driver.sleep(Timeout.reloadWindow);
    await VSBrowser.instance.takeScreenshot(getScreenshotName("before"));
    await ensureExtensionActivated();
  }

  public async after(): Promise<void> {
    await VSBrowser.instance.takeScreenshot(getScreenshotName("after"));
    await this.context?.close();
    await this.browser?.close();
  }

  public async getAadObjectId(envName = "local"): Promise<string> {
    const userDataFile = path.join(
      TestFilePath.configurationFolder,
      `.env.${envName}`
    );
    const configFilePath = path.resolve(
      this.testRootFolder,
      this.appName,
      userDataFile
    );
    const context = dotenvUtil.deserialize(
      await fs.readFile(configFilePath, { encoding: "utf8" })
    );
    const result = context.obj.AAD_APP_OBJECT_ID as string;
    console.log(`TEAMS APP OBJECT ID: ${result}`);
    return result;
  }

  public async getBotAppId(): Promise<string> {
    const userDataFile = path.join(
      TestFilePath.configurationFolder,
      `.env.local`
    );
    const configFilePath = path.resolve(
      this.testRootFolder,
      this.appName,
      userDataFile
    );
    const context = dotenvUtil.deserialize(
      await fs.readFile(configFilePath, { encoding: "utf8" })
    );
    const result = context.obj.BOT_ID as string;
    console.log(`TEAMS BOT ID: ${result}`);
    return result;
  }

  public async getBotObjectId(): Promise<string> {
    const userDataFile = path.join(".fx", "states", "state.local.json");
    const configFilePath = path.resolve(
      this.testRootFolder,
      this.appName,
      userDataFile
    );
    const context = await fs.readJSON(configFilePath);
    const result = context["fx-resource-bot"]["objectId"] as string;
    console.log(`fx-resource-bot.objectId: ${result}`);
    return result;
  }

  public async cleanResource(
    hasAadPlugin = true,
    hasBotPlugin = false,
    envName = "local",
    hasResourceGroup = false
  ): Promise<void> {
    try {
      const cleanService = await GraphApiCleanHelper.create(
        Env.cleanTenantId,
        Env.cleanClientId,
        Env.username,
        Env.password,
        Env.servicePrincipalId,
        Env.servicePrincipalSecret
      );
      if (hasAadPlugin) {
        const aadObjectId = await this.getAadObjectId(envName);
        console.log(`delete AAD ${aadObjectId}`);
        await cleanService.deleteAad(aadObjectId);
      }

      if (hasBotPlugin) {
        const botAppId = await this.getBotAppId();
        const botObjectId = await cleanService.getAadObjectId(botAppId);
        if (botObjectId) {
          console.log(`delete Bot AAD ${botObjectId}`);
          await cleanService.deleteAad(botObjectId);
        }
      }
      if (hasResourceGroup) {
        cleanUpResourceGroup(this.appName, "local");
      }
      await this.cleanDevTunnel();
    } catch (e) {
      console.log(`Failed to clean resource, error message: ${e.message}`);
    }
    await cleanTeamsApp(this.appName);
    await cleanAppStudio(this.appName);
  }

  public async cleanUpResourceGroup(
    appName: string,
    envName?: string
  ): Promise<boolean> {
    if (!appName) {
      return false;
    }
    const name = `${appName}-${envName}-rg`;
    return await deleteResourceGroupByName(name);
  }

  public async cleanDevTunnel(): Promise<void> {
    console.log(`clean dev tunnel`);
    try {
      const devTunnelFilePath = path.resolve(
        this.testRootFolder,
        this.appName,
        "devtunnel.state.json"
      );
      const isExist = await fs.pathExists(devTunnelFilePath);
      if (!isExist) {
        return;
      }
      const devTunnelState = await fs.readJson(
        path.resolve(this.testRootFolder, this.appName, "devtunnel.state.json")
      );
      if (!Array.isArray(devTunnelState?.["teamsToolkit:devtunnel"])) {
        return;
      }
      const devTunnelCleanHelper = await DevTunnelCleanHelper.create(
        Env.cleanTenantId,
        Env.username,
        Env.password,
        Env.servicePrincipalId,
        Env.servicePrincipalSecret
      );
      for (const state of devTunnelState?.["teamsToolkit:devtunnel"]) {
        console.log(`clean tunnel ${state.tunnelId}`);
        await devTunnelCleanHelper.delete(state.tunnelId, state.clusterId);
      }
    } catch (e) {
      console.log(`Failed to clean dev tunnel, error message: ${e.message}`);
    }
  }
}

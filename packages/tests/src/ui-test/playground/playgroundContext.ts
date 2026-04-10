// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs-extra";
import * as path from "path";
import { VSBrowser } from "vscode-extension-tester";
import { TestFilePath } from "../../utils/constants";
import { dotenvUtil } from "../../utils/envUtil";
import { execCommand } from "../../utils/execCommand";
import {
  openExistingProject,
  stopDebugging,
} from "../../utils/vscodeOperation";
import { TestContext } from "../testContext";

export class PlaygroundTestContext extends TestContext {
  public async before(): Promise<void> {
    await super.before();
    await this.createProject();
    await VSBrowser.instance.driver.sleep(30000);

    const testFolder = path.resolve(this.testRootFolder, this.appName);
    await openExistingProject(testFolder);
  }

  public async after(): Promise<void> {
    await stopDebugging();
    await this.context?.close();
    await this.browser?.close();
    await this.cleanResource(false, true);
  }

  public async createProject(): Promise<void> {
    await execCommand(
      this.testRootFolder,
      `atk new --app-name ${this.appName} --interactive false --capability default-bot --programming-language javascript --telemetry false`,
    );
  }

  public async getTeamsAppId(): Promise<string> {
    const userDataFile = path.join(
      TestFilePath.configurationFolder,
      ".env.playground",
    );
    const configFilePath = path.resolve(
      this.testRootFolder,
      this.appName,
      userDataFile,
    );
    const context = dotenvUtil.deserialize(
      await fs.readFile(configFilePath, { encoding: "utf8" }),
    );
    const result = context.obj.TEAMS_APP_ID as string;

    console.log(`TEAMS APP ID: ${result}`);
    return result;
  }
}

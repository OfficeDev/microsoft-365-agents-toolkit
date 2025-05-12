// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { LogProvider } from "@microsoft/teamsfx-api";
import axios, { AxiosInstance } from "axios";
import { exec, execSync } from "child_process";
import * as fs from "fs-extra";
import { glob } from "glob";
import lodash from "lodash";
import os from "os";
import { cpUtils, DebugLogger } from "../../../deps-checker/util/cpUtils";
import { Constants } from "./constants";

export class Utils {
  static async configure(configurePath: string, map: Map<string, string>): Promise<void> {
    let files: string[] = [];
    const extensions = ["*.json", "*.ts", "*.js", "*.scss", "*.tsx"];

    if (fs.lstatSync(configurePath).isFile()) {
      files = [configurePath];
    } else {
      for (const ext of extensions) {
        files = files.concat(glob.sync(`${configurePath}/**/${ext}`, { nodir: true }));
      }
    }

    for (const file of files) {
      let content = (await fs.readFile(file)).toString();
      map.forEach((value, key) => {
        const reg = new RegExp(key, "g");
        content = content.replace(reg, value);
      });
      await fs.writeFile(file, content);
    }
  }

  static normalizeComponentName(name: string): string {
    name = lodash.camelCase(name);
    name = lodash.upperFirst(name);
    return name;
  }

  static async execute(
    command: string,
    title?: string,
    workingDir?: string,
    logProvider?: LogProvider,
    showInOutputWindow = false
  ): Promise<string> {
    return new Promise((resolve, reject) => {
      if (showInOutputWindow) {
        logProvider?.info(`[${title || ""}] Start to run command: "${command}".`);
      }

      exec(command, { cwd: workingDir }, (error, standardOutput) => {
        if (showInOutputWindow) {
          logProvider?.debug(`[${title || ""}]${standardOutput}`);
        }
        if (error) {
          if (showInOutputWindow) {
            logProvider?.error(`[${title || ""}] Failed to run command: "${command}".`);
            logProvider?.error(error.message);
          }
          reject(error);
          return;
        }
        resolve(standardOutput);
      });
    });
  }

  static createAxiosInstanceWithToken(accessToken: string): AxiosInstance {
    const axiosInstance = axios.create({
      headers: {
        authorization: `Bearer ${accessToken}`,
      },
    });
    return axiosInstance;
  }

  static getPackageVersion(pkgName: string): string | undefined {
    try {
      const output = execSync(`npm list ${pkgName} -g --depth=0`);

      const regex = /(?<installPath>[^\n]+)\n`-- ([^@]+)@(?<version>\d+\.\d+\.\d+)/;
      const match = regex.exec(output.toString());
      if (match && match.groups) {
        return match.groups.version;
      } else {
        return undefined;
      }
    } catch (e) {
      return undefined;
    }
  }

  static async hasNPM(logger: DebugLogger | undefined): Promise<boolean> {
    const version = await this.getNPMMajorVersion(logger);
    return version !== undefined;
  }

  static async getNPMMajorVersion(logger: DebugLogger | undefined): Promise<string | undefined> {
    try {
      const output = await cpUtils.executeCommand(
        undefined,
        logger,
        { shell: getShellOptionValue() },
        "npm",
        "--version"
      );

      const regex = /(?<majorVersion>\d+)(\.\d+\.\d+)/;
      const match = regex.exec(output.toString());
      if (match && match.groups) {
        return match.groups.majorVersion;
      } else {
        return undefined;
      }
    } catch (error) {
      return undefined;
    }
  }

  static async getNodeVersion(): Promise<string | undefined> {
    try {
      const output = await cpUtils.executeCommand(
        undefined,
        undefined,
        { shell: getShellOptionValue() },
        "node",
        "--version"
      );

      const regex = /v(?<major_version>\d+)\.(?<minor_version>\d+)\.(?<patch_version>\d+)/gm;
      const match = regex.exec(output);
      if (match && match.groups) {
        return match.groups.major_version;
      } else {
        return undefined;
      }
    } catch (error) {
      return undefined;
    }
  }

  static async findGloballyInstalledVersion(
    logger: LogProvider | undefined,
    packageName: string,
    timeoutInSeconds: number,
    shouldThrowIfNotFound = true
  ): Promise<string | undefined> {
    const timeout = timeoutInSeconds * 1000;
    try {
      const output = await cpUtils.executeCommand(
        undefined,
        logger,
        { timeout: timeout, shell: getShellOptionValue() },
        getExecCommand("npm"),
        "ls",
        `${packageName}`,
        "-g",
        "--depth=0"
      );

      const regex = new RegExp(packageName + "@" + "(?<version>\\d+\\.\\d+\\.\\d+[\\w-.]*)"); // in case user has installed any -alpha, -beta version
      const match = regex.exec(output.toString());
      if (match && match.groups) {
        return match.groups.version;
      } else {
        return undefined;
      }
    } catch (error) {
      logger?.debug(`Failed to execute "npm ls ${packageName}"`);
      if (shouldThrowIfNotFound) {
        throw error;
      }
    }
  }

  static async findLatestVersion(
    logger: LogProvider | undefined,
    packageName: string,
    timeoutInSeconds: number
  ): Promise<string | undefined> {
    const timeout = timeoutInSeconds * 1000;
    try {
      const output = await cpUtils.executeCommand(
        undefined,
        logger,
        { timeout: timeout, shell: getShellOptionValue() },
        getExecCommand("npm"),
        "view",
        `${packageName}`,
        "version"
      );

      const regex = new RegExp("(?<version>\\d+\\.\\d+\\.\\d)");
      const match = regex.exec(output.toString());

      if (match && match.groups) {
        return match.groups.version;
      } else {
        return undefined;
      }
    } catch (error) {
      return undefined;
    }
  }

  static truncateAppShortName(appName: string): string {
    const appNameSuffixPlaceholder = "${{APP_NAME_SUFFIX}}";
    const localSuffix = "local";

    if (appName.endsWith(appNameSuffixPlaceholder)) {
      const appNameWithouSuffix = appName.substring(
        0,
        appName.length - appNameSuffixPlaceholder.length
      );
      if (appNameWithouSuffix.length + localSuffix.length > Constants.TEAMS_APP_NAME_MAX_LENGTH) {
        return (
          appNameWithouSuffix.substring(
            0,
            Constants.TEAMS_APP_NAME_MAX_LENGTH - localSuffix.length
          ) + appNameSuffixPlaceholder
        );
      }
    } else if (appName.length > Constants.TEAMS_APP_NAME_MAX_LENGTH) {
      return appName.substring(0, Constants.TEAMS_APP_NAME_MAX_LENGTH);
    }

    return appName;
  }
}

export function getExecCommand(command: string): string {
  return isWindows() ? `${command}.cmd` : command;
}

export function getShellOptionValue(): boolean | string {
  return isWindows() ? "cmd.exe" : true;
}

function isWindows(): boolean {
  return os.type() === "Windows_NT";
}

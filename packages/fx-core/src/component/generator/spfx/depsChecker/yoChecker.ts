// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ConfigFolderName,
  Context,
  err,
  FxError,
  LogProvider,
  ok,
  Result,
  SystemError,
  UserError,
} from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import * as os from "os";
import * as path from "path";
import { NpmInstallError } from "../../../../error";
import { cpUtils } from "../../../deps-checker/util/cpUtils";
import { DependencyValidateError } from "../error";
import { Constants } from "../utils/constants";
import { telemetryHelper } from "../utils/telemetry-helper";
import { TelemetryEvents, TelemetryProperty } from "../utils/telemetryEvents";
import { getExecCommand, getShellOptionValue, Utils } from "../utils/utils";
import { DependencyChecker } from "./dependencyChecker";

const name = Constants.YeomanPackageName;
const displayName = `${name}`;
const timeout = 6 * 60 * 1000;

export class YoChecker implements DependencyChecker {
  private readonly _logger: LogProvider;

  constructor(logger: LogProvider) {
    this._logger = logger;
  }

  public async ensureDependency(
    ctx: Context,
    targetVersion: string
  ): Promise<Result<boolean, FxError>> {
    telemetryHelper.sendSuccessEvent(ctx, TelemetryEvents.EnsureYoStart);
    try {
      void this._logger.info(`${displayName}@${targetVersion} not found, installing...`);
      await this.install(targetVersion);
      void this._logger.info(`Successfully installed ${displayName}@${targetVersion}`);

      telemetryHelper.sendSuccessEvent(ctx, TelemetryEvents.EnsureYo);
    } catch (error: any) {
      telemetryHelper.sendErrorEvent(
        ctx,
        TelemetryEvents.EnsureYo,
        error as UserError | SystemError,
        {
          [TelemetryProperty.EnsureYoReason]: (error as UserError | SystemError).name,
        }
      );
      this._logger.error(
        `Failed to install ${displayName}@${targetVersion}, error = '${error.toString() as string}'`
      );
      return err(error as UserError | SystemError);
    }

    return ok(true);
  }

  public async findLocalInstalledVersion(): Promise<string | undefined> {
    try {
      const yoVersion = await this.queryVersion();
      const hasSentinel = await fs.pathExists(this.getSentinelPath());
      return hasSentinel ? yoVersion : undefined;
    } catch (error) {
      return undefined;
    }
  }

  public async isLatestInstalled(): Promise<boolean> {
    try {
      const yoVersion = await this.findLocalInstalledVersion();
      const latestYeomanVersion = await this.findLatestVersion(10);
      return !!latestYeomanVersion && yoVersion === latestYeomanVersion;
    } catch (error) {
      return false;
    }
  }

  public async install(targetVersion: string): Promise<void> {
    void this._logger.info("Start installing...");
    await this.cleanup();
    await this.installYo(targetVersion);

    void this._logger.info("Validating package...");
    if (!(await this.validate())) {
      void this._logger.debug("Failed to validate yo, cleaning up...");
      await this.cleanup();
      throw DependencyValidateError(name);
    }
  }

  public getBinFolders(): string[] {
    const defaultPath = this.getDefaultInstallPath();
    return [defaultPath, path.join(defaultPath, "node_modules", ".bin")];
  }

  public async findGloballyInstalledVersion(
    timeoutInSeconds?: number
  ): Promise<string | undefined> {
    return await Utils.findGloballyInstalledVersion(this._logger, name, timeoutInSeconds ?? 0);
  }

  public async findLatestVersion(timeoutInSeconds: number): Promise<string | undefined> {
    return await Utils.findLatestVersion(this._logger, name, timeoutInSeconds);
  }

  private async validate(): Promise<boolean> {
    return await fs.pathExists(this.getSentinelPath());
  }

  private getDefaultInstallPath(): string {
    return path.join(os.homedir(), `.${ConfigFolderName}`, "bin", "yo");
  }

  private getSentinelPath(): string {
    return path.join(os.homedir(), `.${ConfigFolderName}`, "yo-sentinel");
  }

  private async queryVersion(): Promise<string | undefined> {
    const packagePath = path.join(
      this.getDefaultInstallPath(),
      "node_modules",
      "yo",
      "package.json"
    );
    if (await fs.pathExists(packagePath)) {
      const packageJson = await fs.readJson(packagePath);
      return packageJson.version ?? undefined;
    }
    return undefined;
  }

  private async cleanup(): Promise<void> {
    try {
      const legacyDirectory = path.join(os.homedir(), `.${ConfigFolderName}`, "bin", "spfx");
      if (fs.existsSync(legacyDirectory)) {
        await fs.emptyDir(legacyDirectory);
        await fs.rmdir(legacyDirectory);
      }

      await fs.emptyDir(this.getDefaultInstallPath());
      await fs.remove(this.getSentinelPath());

      const yoExecutables = [
        "yo",
        "yo.cmd",
        "yo.ps1",
        "yo-complete",
        "yo-complete.cmd",
        "yo-complete.ps1",
      ];
      await Promise.all(
        yoExecutables.map(async (executable) => {
          const executablePath = path.join(this.getDefaultInstallPath(), executable);
          if (await fs.pathExists(executablePath)) {
            await fs.remove(executablePath);
          }
        })
      );
    } catch (err: any) {
      this._logger.error(
        `Failed to clean up path: ${this.getDefaultInstallPath()}, error: ${
          err.toString() as string
        }`
      );
    }
  }

  private async installYo(targetVersion: string): Promise<void> {
    const version = targetVersion ?? Constants.LatestVersion;
    try {
      await fs.ensureDir(path.join(this.getDefaultInstallPath(), "node_modules"));
      await cpUtils.executeCommand(
        undefined,
        this._logger,
        { timeout: timeout, shell: getShellOptionValue() },
        getExecCommand("npm"),
        "install",
        `${name}@${version}`,
        "--prefix",
        `${this.getDefaultInstallPath()}`,
        "--no-audit",
        "--global-style"
      );

      await fs.ensureFile(this.getSentinelPath());
    } catch (error) {
      void this._logger.error(`Failed to execute npm install ${displayName}@${version}`);
      throw new NpmInstallError(error as Error);
    }
  }
}

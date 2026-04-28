// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as fs from "fs-extra";
import * as os from "os";
import * as path from "path";

import { ConfigFolderName, UserError } from "@microsoft/teamsfx-api";
import {
  VxTestAppInvalidInstallOptionsError,
  VxTestAppValidationError,
} from "../../../error/depCheck";
import { BaseInstallOptions, DependencyStatus, DepsChecker, DepsType } from "../depsChecker";
import { DepsLogger } from "../depsLogger";
import { DepsTelemetry } from "../depsTelemetry";
import { isMacOS, isWindows } from "../util";
import { downloadToTempFile, unzip } from "../util/downloadHelper";
import { createSymlink } from "../util/fileHelper";

interface InstallOptionsSafe {
  version: string;
  projectPath: string;
}

const VxTestAppName = "Video Extensibility Test App";

// https://www.electronjs.org/docs/latest/tutorial/application-distribution#manual-packaging
const VxTestAppExecutableName = isWindows()
  ? "video-extensibility-test-app.exe"
  : isMacOS()
    ? "video-extensibility-test-app.app"
    : "video-extensibility-test-app";

const VxTestAppDirRelPath = path.join(".tools", "video-extensibility-test-app");
const VxTestAppGlobalBasePath = path.join(
  os.homedir(),
  `.${ConfigFolderName}`,
  `bin`,
  `video-extensibility-test-app`
);
const VxTestAppDownloadTimeoutMillis = 5 * 60 * 1000;
// TODO: change to GitHub release after new VxTestApp is released.
const VxTestAppDownloadUrlTemplate =
  "https://github.com/microsoft/teams-videoapp-sample/releases/download/testApp-v@version/video-extensibility-test-app-@platform-@arch-portable.zip";

export class VxTestAppChecker implements DepsChecker {
  private readonly _logger: DepsLogger;
  private readonly _telemetry: DepsTelemetry;

  constructor(logger: DepsLogger, telemetry: DepsTelemetry) {
    this._logger = logger;
    this._telemetry = telemetry;
  }

  public async resolve(installOptions?: BaseInstallOptions): Promise<DependencyStatus> {
    if (!this.isValidInstallOptions(installOptions)) {
      return VxTestAppChecker.newDependencyStatusForInstallError(
        // documentation no longer exists, replaced with empty string.
        new VxTestAppInvalidInstallOptionsError()
      );
    }

    // check installation in project dir
    const installInfo = await this.getInstallationInfo(installOptions);
    if (installInfo.isInstalled) {
      return installInfo;
    }

    // ensure vxTestApp is installed in global dir
    const globalInstallDir = path.join(VxTestAppGlobalBasePath, installOptions.version);
    if (!(await this.isValidInstalltion(globalInstallDir, installOptions.version))) {
      await fs.remove(globalInstallDir);
      await this.installVersion(installOptions.version, globalInstallDir);
    }

    // ensure vxTestApp is installed in project dir
    const projectInstallDir = path.join(installOptions.projectPath, VxTestAppDirRelPath);
    await createSymlink(globalInstallDir, projectInstallDir);
    // TODO: need to chmod to add executable permission for non-Windows OS
    if (!(await this.isValidInstalltion(projectInstallDir, installOptions.version))) {
      return VxTestAppChecker.newDependencyStatusForInstallError(
        // documentation no longer exists, replaced with empty string.
        new VxTestAppValidationError()
      );
    }

    return {
      name: VxTestAppName,
      type: DepsType.VxTestApp,
      isInstalled: true,
      command: VxTestAppExecutableName,
      details: {
        isLinuxSupported: false,
        supportedVersions: [installOptions.version],
        binFolders: [projectInstallDir],
      },
      error: undefined,
    };
  }

  public async getInstallationInfo(installOptions?: BaseInstallOptions): Promise<DependencyStatus> {
    if (!this.isValidInstallOptions(installOptions)) {
      return VxTestAppChecker.newDependencyStatusForInstallError(
        // documentation no longer exists, replaced with empty string.
        new VxTestAppValidationError()
      );
    }

    const installDir = path.join(installOptions.projectPath, VxTestAppDirRelPath);
    if (!(await this.isValidInstalltion(installDir, installOptions.version))) {
      return VxTestAppChecker.newDependencyStatusForNotInstalled(installOptions.version);
    }

    const projectInstallDir = path.join(installOptions.projectPath, VxTestAppDirRelPath);
    return {
      name: VxTestAppName,
      type: DepsType.VxTestApp,
      isInstalled: true,
      command: VxTestAppExecutableName,
      details: {
        isLinuxSupported: false,
        supportedVersions: [installOptions.version],
        binFolders: [projectInstallDir],
      },
      error: undefined,
    };
  }

  private async installVersion(version: string, installDir: string): Promise<void> {
    const downloadUrl = this.getDownloadUrl(version);
    await downloadToTempFile(
      downloadUrl,
      { timeout: VxTestAppDownloadTimeoutMillis },
      async (zipFilePath: string) => {
        await unzip(zipFilePath, installDir);
      }
    );
  }

  private getDownloadUrl(version: string): string {
    const url = VxTestAppDownloadUrlTemplate.replace(/@version/g, version)
      .replace(/@platform/g, os.platform())
      .replace(/@arch/g, os.arch());

    return url;
  }

  private async isValidInstalltion(installDir: string, version: string): Promise<boolean> {
    const vxTestAppExecutable = path.join(installDir, VxTestAppExecutableName);
    if (!(await fs.pathExists(vxTestAppExecutable))) {
      return false;
    }

    // TODO(aochengwang):
    //   1. check executable permission for non-Windows OS
    //   2. check whether installed version is correct?
    return true;
  }

  private isValidInstallOptions(
    installOptions?: BaseInstallOptions
  ): installOptions is InstallOptionsSafe {
    return !(installOptions?.projectPath === undefined && installOptions?.version === undefined);
  }

  private static newDependencyStatusForNotInstalled(version?: string): DependencyStatus {
    return {
      name: VxTestAppName,
      type: DepsType.VxTestApp,
      isInstalled: false,
      command: VxTestAppExecutableName,
      details: {
        isLinuxSupported: false,
        supportedVersions: version === undefined ? [] : [version],
      },
      error: undefined,
    };
  }

  private static newDependencyStatusForInstallError(
    error: UserError,
    version?: string
  ): DependencyStatus {
    return {
      name: VxTestAppName,
      type: DepsType.VxTestApp,
      isInstalled: false,
      command: VxTestAppExecutableName,
      details: {
        isLinuxSupported: false,
        supportedVersions: version === undefined ? [] : [version],
      },
      error: error,
    };
  }
}

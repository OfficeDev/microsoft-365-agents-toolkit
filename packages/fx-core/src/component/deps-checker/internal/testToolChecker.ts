// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ConfigFolderName, err, ok, Result, UserError } from "@microsoft/teamsfx-api";
import * as fs from "fs-extra";
import fetch from "node-fetch";
import * as os from "os";
import * as path from "path";
import semver from "semver";
import maxSatisfying from "semver/ranges/max-satisfying";
import * as url from "url";
import * as uuid from "uuid";
import { getLocalizedString } from "../../../common/localizeUtils";
import { DepsCheckerError, NodejsNotFoundError } from "../../../error";
import { v3DefaultHelpLink } from "../constant/helpLink";
import { Messages } from "../constant/message";
import { TelemetryProperties } from "../constant/telemetry";
import {
  DependencyStatus,
  DepsChecker,
  DepsType,
  TestToolInstallOptions,
  TestToolReleaseType,
} from "../depsChecker";
import { cpUtils } from "../util";
import { downloadToTempFile, unzip } from "../util/downloadHelper";
import { cleanup, createSymlink, rename, unlinkSymlink } from "../util/fileHelper";
import { isWindows } from "../util/system";

enum InstallType {
  Global = "global",
  Portable = "portable",
}

type TestToolDependencyStatus = Omit<DependencyStatus, "isInstalled"> &
  ({ isInstalled: true; installType: InstallType } | { isInstalled: false });

interface InstallationInfoFile {
  lastCheckTimestamp: number;
}

const InstallTimeout = 5 * 60 * 1000;

export class TestToolChecker implements DepsChecker {
  private telemetryProperties: { [key: string]: string };
  private readonly name = "Microsoft 365 Agents Playground";
  private readonly npmPackageName = "@microsoft/teams-app-test-tool";
  private readonly checkUpdateTimeout = 10 * 1000;
  private readonly npmCommandName = isWindows() ? "teamsapptester.cmd" : "teamsapptester";
  private readonly binaryCommandName = isWindows() ? "teamsapptester.exe" : "teamsapptester";
  private readonly portableDirNameNpm = "testTool";
  private readonly portableDirNameBinary = "testToolBinary";
  // Limit 1 hour check update internval because of GitHub API throttling limitation
  // https://docs.github.com/en/rest/using-the-rest-api/rate-limits-for-the-rest-api?apiVersion=2022-11-28#primary-rate-limit-for-unauthenticated-users
  private readonly defaultUpdateInterval = 1 * 60 * 60 * 1000;

  constructor() {
    this.telemetryProperties = {};
  }

  public async getInstallationInfo(
    installOptions: TestToolInstallOptions
  ): Promise<TestToolDependencyStatus> {
    this.telemetryProperties[TelemetryProperties.InstallTestToolReleaseType] =
      installOptions.releaseType;

    const symlinkDir = installOptions.symlinkDir
      ? path.resolve(installOptions.projectPath, installOptions.symlinkDir)
      : undefined;

    // check version in symlink dir
    if (symlinkDir) {
      const versionRes = await this.checkVersion(
        installOptions.releaseType,
        installOptions.versionRange,
        this.getBinFolder(installOptions.releaseType, symlinkDir)
      );
      if (versionRes.isOk()) {
        this.telemetryProperties[TelemetryProperties.SymlinkTestToolVersion] = versionRes.value;
        return await this.getSuccessDepsInfo(versionRes.value, symlinkDir);
      } else {
        this.telemetryProperties[TelemetryProperties.SymlinkTestToolVersionError] =
          versionRes.error.message;
        await unlinkSymlink(symlinkDir);
      }
    }

    // check version in ${HOME}/.fx/bin and createSymlink if asked for
    const version = await this.findLatestInstalledPortableVersion(
      installOptions.releaseType,
      installOptions.versionRange
    );
    if (version) {
      const portablePath = path.join(
        this.getPortableVersionsDir(installOptions.releaseType),
        version
      );
      this.telemetryProperties[TelemetryProperties.SelectedPortableTestToolVersion] = version;
      if (symlinkDir) {
        await createSymlink(portablePath, symlinkDir);
        return await this.getSuccessDepsInfo(version, symlinkDir);
      } else {
        return await this.getSuccessDepsInfo(version, portablePath);
      }
    }

    // check global version in PATH
    const globalVersionRes = await this.checkVersion(
      installOptions.releaseType,
      installOptions.versionRange
    );
    if (globalVersionRes.isOk()) {
      const version = globalVersionRes.value;
      this.telemetryProperties[TelemetryProperties.GlobalTestToolVersion] = version;
      return this.getSuccessDepsInfo(version, undefined);
    } else {
      this.telemetryProperties[TelemetryProperties.GlobalTestToolVersionError] =
        globalVersionRes.error.message;
    }

    return this.createFailureDepsInfo(installOptions.versionRange, undefined);
  }

  public async resolve(installOptions: TestToolInstallOptions): Promise<DependencyStatus> {
    this.telemetryProperties[TelemetryProperties.InstallTestToolReleaseType] =
      installOptions.releaseType;

    let installationInfo: TestToolDependencyStatus;
    try {
      if (installOptions.releaseType === TestToolReleaseType.Npm && !(await this.hasNode())) {
        throw new NodejsNotFoundError();
      }
      installationInfo = await this.getInstallationInfo(installOptions);
      if (!installationInfo.isInstalled) {
        const symlinkDir = installOptions.symlinkDir
          ? path.resolve(installOptions.projectPath, installOptions.symlinkDir)
          : undefined;
        installationInfo = await this.install(
          installOptions.releaseType,
          installOptions.projectPath,
          installOptions.versionRange,
          symlinkDir
        );
      } else {
        if (installationInfo.installType === InstallType.Portable) {
          const updateInstallationInfo = await this.autoUpdate(installOptions);
          if (updateInstallationInfo) {
            installationInfo = updateInstallationInfo;
          }
        }
      }

      return installationInfo;
    } catch (error: any) {
      if (error instanceof UserError) {
        return await this.createFailureDepsInfo(installOptions.versionRange, error);
      }
      return await this.createFailureDepsInfo(
        installOptions.versionRange,
        new DepsCheckerError(error.message, v3DefaultHelpLink)
      );
    }
  }

  private async install(
    releaseType: TestToolReleaseType,
    projectPath: string,
    versionRange: string,
    symlinkDir?: string
  ): Promise<TestToolDependencyStatus> {
    if (releaseType === TestToolReleaseType.Npm && !(await this.hasNPM())) {
      throw new DepsCheckerError(Messages.needInstallNpm(), v3DefaultHelpLink);
    }

    const tmpVersion = `tmp-${uuid.v4().slice(0, 6)}`;
    const tmpPath = this.getPortableInstallPath(releaseType, tmpVersion);
    await fs.ensureDir(tmpPath);
    if (releaseType === TestToolReleaseType.Npm) {
      await this.npmInstall(projectPath, tmpPath, versionRange);
    } else {
      await this.binaryInstall(tmpPath, versionRange);
    }
    const versionRes = await this.checkVersion(
      releaseType,
      versionRange,
      this.getBinFolder(releaseType, tmpPath)
    );
    if (versionRes.isErr()) {
      await cleanup(tmpPath);
      this.telemetryProperties[TelemetryProperties.InstallTestToolError] = versionRes.error.message;
      throw new DepsCheckerError(
        Messages.failToValidateTestTool(versionRes.error.message),
        v3DefaultHelpLink
      );
    }
    const actualVersion = versionRes.value;
    this.telemetryProperties[TelemetryProperties.InstalledTestToolVersion] = actualVersion;

    const portablePath = this.getPortableInstallPath(releaseType, actualVersion);
    await rename(tmpPath, portablePath);

    if (symlinkDir) {
      await createSymlink(portablePath, symlinkDir);
    }

    await this.writeInstallInfoFile(projectPath);

    if (symlinkDir) {
      return await this.getSuccessDepsInfo(actualVersion, symlinkDir);
    } else {
      return await this.getSuccessDepsInfo(actualVersion, portablePath);
    }
  }

  private async hasNewVersionReleasedInRange(
    releaseType: TestToolReleaseType,
    latestInstalledVersion: string,
    versionRange: string
  ): Promise<boolean> {
    if (releaseType === TestToolReleaseType.Npm) {
      try {
        const result = await cpUtils.executeCommand(
          undefined,
          undefined,
          // avoid powershell execution policy issue.
          { shell: isWindows() ? "cmd.exe" : true, timeout: this.checkUpdateTimeout },
          "npm",
          "view",
          `"${this.npmPackageName}@${versionRange}"`,
          "version",
          "--json"
        );
        // when there are one result, it will return string
        // when there are multiple results, it will return array of strings
        let versionList: string[] | string = JSON.parse(result);
        if (typeof versionList === "string") {
          versionList = [versionList];
        }
        if (!Array.isArray(versionList)) {
          // do update if npm returned invalid result
          return true;
        }
        return versionList.filter((v) => semver.gt(v, latestInstalledVersion)).length > 0;
      } catch {
        // just a best effort optimization to save one download if no recent version has been released
        // do update if check failed
        return true;
      }
    } else {
      // get version list
      const releases = await GitHubHelpers.listGitHubReleases();
      const versionList = releases.map((release) => release.version);
      return versionList.filter((v) => semver.gt(v, latestInstalledVersion)).length > 0;
    }
  }

  // return undefined if not updated or update failure
  private async autoUpdate(
    installOptions: TestToolInstallOptions
  ): Promise<TestToolDependencyStatus | undefined> {
    const installInfo = await this.readInstallInfoFile(installOptions.projectPath);
    const now = new Date().getTime();
    const updateExpired =
      !installInfo || now > installInfo.lastCheckTimestamp + this.defaultUpdateInterval;

    if (!updateExpired) {
      return undefined;
    }

    const latestInstalledVersion = await this.findLatestInstalledPortableVersion(
      installOptions.releaseType,
      installOptions.versionRange
    );
    if (
      latestInstalledVersion !== undefined &&
      !(await this.hasNewVersionReleasedInRange(
        installOptions.releaseType,
        latestInstalledVersion,
        installOptions.versionRange
      ))
    ) {
      return undefined;
    }

    this.telemetryProperties[TelemetryProperties.TestToolLastUpdateTimestamp] =
      installInfo?.lastCheckTimestamp?.toString() || "<never>";
    this.telemetryProperties[TelemetryProperties.TestToolUpdatePreviousVersion] =
      latestInstalledVersion || "<undefined>";
    const symlinkDir = installOptions.symlinkDir
      ? path.resolve(installOptions.projectPath, installOptions.symlinkDir)
      : undefined;

    try {
      return await this.install(
        installOptions.releaseType,
        installOptions.projectPath,
        installOptions.versionRange,
        symlinkDir
      );
    } catch (e: unknown) {
      // ignore update failure and use existing version
      if (e instanceof Error) {
        this.telemetryProperties[TelemetryProperties.TestToolUpdateError] = e.message;
      }
      await this.writeInstallInfoFile(installOptions.projectPath);
      return undefined;
    }
  }

  private validateInstallInfoFile(data: unknown): data is InstallationInfoFile {
    if ("lastCheckTimestamp" in (data as InstallationInfoFile)) {
      if (typeof (data as InstallationInfoFile).lastCheckTimestamp === "number") {
        return true;
      }
    }

    return false;
  }

  private async readInstallInfoFile(
    projectPath: string
  ): Promise<InstallationInfoFile | undefined> {
    const installInfoPath = this.getInstallInfoPath(projectPath);
    try {
      const data: unknown = await fs.readJson(installInfoPath);
      if (this.validateInstallInfoFile(data)) {
        return data;
      }
    } catch {
      // ignore invalid installation info file
    }
    await cleanup(installInfoPath);
    return undefined;
  }

  private async writeInstallInfoFile(projectPath: string) {
    const projectInfoPath = this.getInstallInfoPath(projectPath);
    const installInfo: InstallationInfoFile = {
      lastCheckTimestamp: new Date().getTime(),
    };
    await fs.ensureDir(path.dirname(projectInfoPath));
    await fs.writeJson(projectInfoPath, installInfo);
  }

  private async findLatestInstalledPortableVersion(
    releaseType: TestToolReleaseType,
    versionRange: string
  ): Promise<string | undefined> {
    let portablePath: string | undefined;
    try {
      const portableVersionsDir = this.getPortableVersionsDir(releaseType);
      const dirs = await fs.readdir(portableVersionsDir, { withFileTypes: true });
      const satisfiedVersions = dirs
        .filter(
          (dir) =>
            dir.isDirectory() && semver.valid(dir.name) && semver.satisfies(dir.name, versionRange)
        )
        .map((dir) => dir.name);

      // sort by version desc
      satisfiedVersions.sort((a, b) => semver.rcompare(a, b));

      // find the latest version that is working
      for (const version of satisfiedVersions) {
        portablePath = path.join(portableVersionsDir, version);
        const checkVersionRes = await this.checkVersion(
          releaseType,
          versionRange,
          this.getBinFolder(releaseType, portablePath)
        );
        if (checkVersionRes.isOk()) {
          return version;
        }
        this.telemetryProperties[TelemetryProperties.VersioningTestToolVersionError] =
          (this.telemetryProperties[TelemetryProperties.VersioningTestToolVersionError] ?? "") +
          `[${version}] ${checkVersionRes.error.message}`;
      }
    } catch {
      // ignore errors if portable dir doesn't exist
    }
    return undefined;
  }

  private async checkVersion(
    releaseType: TestToolReleaseType,
    versionRange: string,
    binFolder?: string
  ): Promise<Result<string, DepsCheckerError>> {
    try {
      const actualVersion = await this.queryVersion(releaseType, binFolder);
      if (semver.satisfies(actualVersion, versionRange)) {
        return ok(actualVersion);
      } else {
        return err(
          new DepsCheckerError(
            Messages.testToolVersionNotMatch(actualVersion, versionRange),
            v3DefaultHelpLink
          )
        );
      }
    } catch (error: any) {
      return err(new DepsCheckerError(error.message, v3DefaultHelpLink));
    }
  }

  private async queryVersion(
    releaseType: TestToolReleaseType,
    binFolder: string | undefined
  ): Promise<string> {
    const commandName =
      releaseType === TestToolReleaseType.Npm ? this.npmCommandName : this.binaryCommandName;
    const execPath = binFolder ? path.resolve(binFolder, commandName) : commandName;
    const output = await cpUtils.executeCommand(
      undefined,
      undefined,
      // avoid powershell execution policy issue.
      { shell: isWindows() ? "cmd.exe" : true, timeout: InstallTimeout },
      `"${execPath}"`,
      "--version"
    );
    return output.trim();
  }

  private async hasNode(): Promise<boolean> {
    try {
      await cpUtils.executeCommand(undefined, undefined, { shell: true }, "node", "--version");
      return true;
    } catch (error) {
      return false;
    }
  }

  private async hasNPM(): Promise<boolean> {
    try {
      await cpUtils.executeCommand(undefined, undefined, { shell: true }, "npm", "--version");
      return true;
    } catch (error) {
      return false;
    }
  }

  private async npmInstall(
    projectPath: string,
    prefix: string,
    versionRange: string
  ): Promise<void> {
    let pkg: string | undefined;

    // Before we release package to npm registry we can use tgz to test
    pkg ||= await this.findLocalNpmPackage(projectPath);
    pkg ||= await this.findLocalNpmPackage(path.join(projectPath, "devTools"));

    pkg ||= `${this.npmPackageName}@${versionRange}`;

    try {
      await cpUtils.executeCommand(
        undefined,
        undefined,
        // avoid powershell execution policy issue.
        { shell: isWindows() ? "cmd.exe" : true, timeout: InstallTimeout },
        `npm`,
        "install",
        `"${pkg}"`,
        "--prefix",
        `"${prefix}"`,
        "--no-audit"
      );
    } catch (error: any) {
      await cleanup(prefix);
      // @ is incorrectly identified as an email format.
      this.telemetryProperties[TelemetryProperties.InstallTestToolError] = (error.message as string)
        ?.split(pkg)
        ?.join(pkg);
      throw new DepsCheckerError(
        getLocalizedString("error.common.InstallSoftwareError", this.name),
        v3DefaultHelpLink
      );
    }
  }

  // TODO: remove after release to npm
  private async findLocalNpmPackage(dir: string): Promise<string | undefined> {
    try {
      const files = await fs.readdir(dir);
      for (const fileName of files) {
        const fullPath = path.join(dir, fileName);
        if (fileName.match(/microsoft-teams-app-test-tool.*\.tgz/i)) {
          try {
            const st = await fs.stat(fullPath);
            if (st.isFile()) {
              // encode special characters in path
              return url.pathToFileURL(fullPath).toString();
            }
          } catch {
            // ignore invalid files
          }
        }
      }
    } catch {
      // local npm package error should not block
    }
    return undefined;
  }

  private async binaryInstall(installPath: string, versionRange: string): Promise<void> {
    const releases = await GitHubHelpers.listGitHubReleases();
    const targetVersion = maxSatisfying(
      releases.map((release) => release.version),
      versionRange
    );
    if (targetVersion === null) {
      throw new DepsCheckerError(
        getLocalizedString("error.common.VersionError", versionRange),
        v3DefaultHelpLink
      );
    }
    const release = releases.find((value) => value.version === targetVersion) as GitHubRelease;

    await downloadToTempFile(
      release.url,
      {
        timeout: InstallTimeout,
        headers: {
          Accept: "application/octet-stream",
        },
      },
      async (filePath: string) => {
        await unzip(filePath, installPath);
      }
    );
  }

  private getBinFolder(releaseType: TestToolReleaseType, installPath: string) {
    if (releaseType === TestToolReleaseType.Npm) {
      return path.join(installPath, "node_modules", ".bin");
    } else {
      return path.join(installPath);
    }
  }
  private getPortableVersionsDir(releaseType: string): string {
    if (releaseType === TestToolReleaseType.Npm) {
      return path.join(os.homedir(), `.${ConfigFolderName}`, "bin", this.portableDirNameNpm);
    } else {
      return path.join(os.homedir(), `.${ConfigFolderName}`, "bin", this.portableDirNameBinary);
    }
  }
  private getPortableInstallPath(releaseType: TestToolReleaseType, version: string): string {
    return path.join(this.getPortableVersionsDir(releaseType), version);
  }
  private getInstallInfoPath(projectDir: string): string {
    return path.join(projectDir, "devTools", ".testTool.installInfo.json");
  }
  private async getSuccessDepsInfo(
    version: string,
    binFolder?: string
  ): Promise<TestToolDependencyStatus> {
    return Promise.resolve({
      name: this.name,
      type: DepsType.TestTool,
      isInstalled: true,
      command: this.npmCommandName,
      details: {
        isLinuxSupported: true,
        supportedVersions: [], // unused
        binFolders: binFolder ? [binFolder] : [],
        installVersion: version,
      },
      telemetryProperties: this.telemetryProperties,
      error: undefined,
      installType: binFolder ? InstallType.Portable : InstallType.Global,
    });
  }
  private async createFailureDepsInfo(
    version: string,
    error?: DepsCheckerError
  ): Promise<TestToolDependencyStatus> {
    return Promise.resolve({
      name: this.name,
      type: DepsType.TestTool,
      isInstalled: false,
      command: this.npmCommandName,
      details: {
        isLinuxSupported: true,
        supportedVersions: [], // unused
        binFolders: [],
        installVersion: version,
      },
      telemetryProperties: this.telemetryProperties,
      error: error,
    });
  }
}

export interface GitHubRelease {
  version: string;
  url: string;
}

export class GitHubHelpers {
  private static readonly releasePackageName = "teams-app-test-tool";
  private static readonly artifactNamePrefix = "teamsapptester";
  public static async listGitHubReleases(): Promise<GitHubRelease[]> {
    // GitHub API without auth
    const response = await fetch("https://api.github.com/repos/OfficeDev/TeamsFx/releases", {
      headers: {
        Accept: "application/vnd.github+json",
        "X-Github-Api-Version": "2022-11-28",
      },
      timeout: InstallTimeout,
    });
    const releases: {
      tag_name: string;
      assets: { name: string; url: string }[];
    }[] = await response.json();

    const result: GitHubRelease[] = [];
    for (const release of releases) {
      const parts = release.tag_name.split("@");
      const assets = release.assets.filter((asset) =>
        asset.name.includes(`${this.artifactNamePrefix}-${os.platform()}-${os.arch()}`)
      );
      if (parts.length === 2) {
        const pkgName = parts[0];
        const version = parts[1];
        if (pkgName == this.releasePackageName && assets.length > 0) {
          result.push({
            version,
            url: release.assets[0].url,
          });
        }
      }
    }

    return result;
  }
}

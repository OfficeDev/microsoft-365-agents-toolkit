// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, UserError } from "@microsoft/teamsfx-api";
import * as fs from "fs-extra";
import * as path from "path";
import semver from "semver";
import {
  DepsCheckerError,
  NodejsNotFoundError,
  NodejsNotLtsError,
  NodejsNotRecommendedError,
} from "../../../error";
import {
  nodeNotFoundHelpLink,
  v3NodeNotFoundHelpLink,
  v3NodeNotSupportedHelpLink,
} from "../constant/helpLink";
import { DepsCheckerEvent } from "../constant/telemetry";
import { BaseInstallOptions, DependencyStatus, DepsChecker, DepsType } from "../depsChecker";
import { DepsLogger } from "../depsLogger";
import { DepsTelemetry } from "../depsTelemetry";
import { cpUtils } from "../util/cpUtils";

const NodeName = "Node.js";

class NodeVersion {
  public readonly version: string;
  public readonly majorVersion: string;

  constructor(version: string, majorVersion: string) {
    this.version = version;
    this.majorVersion = majorVersion;
  }
}

export abstract class NodeChecker implements DepsChecker {
  protected abstract readonly _nodeNotFoundHelpLink: string;
  protected abstract readonly _nodeNotSupportedEvent: DepsCheckerEvent;
  protected abstract readonly _type: DepsType;
  protected abstract getSupportedVersions(projectPath?: string): Promise<string[]>;
  protected abstract isVersionSupported(supportedVersions: string[], version: NodeVersion): boolean;
  protected abstract getVersionNotSupportedError(
    supportedVersions: string[],
    version: NodeVersion
  ): UserError;
  protected abstract readonly _minErrorVersion: number;
  protected abstract readonly _maxErrorVersion: number;

  private readonly _telemetry: DepsTelemetry;
  private readonly _logger: DepsLogger;

  constructor(logger: DepsLogger, telemetry: DepsTelemetry) {
    this._logger = logger;
    this._telemetry = telemetry;
  }

  public async getInstallationInfo(installOptions?: BaseInstallOptions): Promise<DependencyStatus> {
    let supportedVersions: string[] = [];
    try {
      supportedVersions = await this.getSupportedVersions(installOptions?.projectPath);

      this._logger.debug(
        `NodeChecker checking for supported versions: '${JSON.stringify(supportedVersions)}'`
      );

      const currentVersion = await NodeChecker.getInstalledNodeVersion();
      if (currentVersion === null) {
        this._telemetry.sendUserErrorEvent(
          DepsCheckerEvent.nodeNotFound,
          "Node.js can't be found."
        );
        const error = new NodejsNotFoundError();
        return await this.getDepsInfo(false, supportedVersions, undefined, error);
      }
      this._telemetry.sendEvent(DepsCheckerEvent.nodeVersion, {
        "global-version": `${currentVersion.version}`,
        "global-major-version": `${currentVersion.majorVersion}`,
      });

      if (!this.isVersionSupported(supportedVersions, currentVersion)) {
        this._telemetry.sendUserErrorEvent(
          this._nodeNotSupportedEvent,
          `Node.js ${currentVersion.version} is not supported.`
        );
        return await this.getDepsInfo(
          !NodeChecker.isVersionError(this._minErrorVersion, this._maxErrorVersion, currentVersion),
          supportedVersions,
          currentVersion.version,
          this.getVersionNotSupportedError(supportedVersions, currentVersion)
        );
      }
      return await this.getDepsInfo(true, supportedVersions, currentVersion.version);
    } catch (error) {
      return await this.getDepsInfo(
        false,
        supportedVersions,
        undefined,
        new DepsCheckerError(error.message, nodeNotFoundHelpLink)
      );
    }
  }

  public async resolve(installOptions?: BaseInstallOptions): Promise<DependencyStatus> {
    const installationInfo = await this.getInstallationInfo(installOptions);
    if (installationInfo.error) {
      this._logger.printDetailLog();
      this._logger.error(
        `${installationInfo.error.message}, error = '${installationInfo.error.toString()}'`
      );
    }
    this._logger.cleanup();
    return installationInfo;
  }

  public async install(): Promise<void> {
    return Promise.resolve();
  }

  public async getDepsInfo(
    isInstalled: boolean,
    supportedVersions: string[],
    installVersion?: string,
    error?: FxError
  ): Promise<DependencyStatus> {
    return {
      name: NodeName,
      type: this._type,
      isInstalled: isInstalled,
      command: await this.command(),
      details: {
        isLinuxSupported: true,
        supportedVersions: supportedVersions,
        installVersion: installVersion,
      },
      error: error,
    };
  }

  private static isVersionError(
    minErrorVersion: number,
    maxErrorVersion: number,
    version: NodeVersion
  ): boolean {
    const majorVersion = Number.parseInt(version.majorVersion);

    return (
      !Number.isInteger(majorVersion) ||
      majorVersion <= minErrorVersion ||
      majorVersion >= maxErrorVersion
    );
  }

  public command(): Promise<string> {
    return Promise.resolve("node");
  }

  public static async getInstalledNodeVersion(): Promise<NodeVersion | null> {
    try {
      const output = await cpUtils.executeCommand(
        undefined,
        undefined,
        undefined,
        "node",
        "--version"
      );
      return getNodeVersion(output);
    } catch (error) {
      return null;
    }
  }
}

function getNodeVersion(output: string): NodeVersion | null {
  const regex = /v(?<major_version>\d+)\.(?<minor_version>\d+)\.(?<patch_version>\d+)/gm;
  const match = regex.exec(output);
  if (!match) {
    return null;
  }

  const majorVersion = match.groups?.major_version;
  if (!majorVersion) {
    return null;
  }

  return new NodeVersion(match[0], majorVersion);
}

export class LtsNodeChecker extends NodeChecker {
  protected readonly _nodeNotFoundHelpLink = v3NodeNotFoundHelpLink;
  protected readonly _nodeNotSupportedEvent = DepsCheckerEvent.nodeNotLts;
  protected readonly _type = DepsType.LtsNode;
  protected readonly _minErrorVersion = Number.MIN_SAFE_INTEGER;
  protected readonly _maxErrorVersion = Number.MAX_SAFE_INTEGER;

  protected getSupportedVersions(): Promise<string[]> {
    return Promise.resolve(["18", "20", "22"]);
  }

  protected isVersionSupported(supportedVersions: string[], version: NodeVersion): boolean {
    return supportedVersions.includes(version.majorVersion);
  }
  protected getVersionNotSupportedError(
    supportedVersions: string[],
    version: NodeVersion
  ): UserError {
    const supportedVersionsString = supportedVersions.map((v) => "v" + v).join(", ");
    return new NodejsNotLtsError(version.version, supportedVersionsString);
  }
}

export class ProjectNodeChecker extends NodeChecker {
  protected readonly _nodeNotFoundHelpLink = v3NodeNotFoundHelpLink;
  protected readonly _nodeNotSupportedEvent = DepsCheckerEvent.nodeNotSupportedForProject;
  protected readonly _type = DepsType.ProjectNode;
  protected readonly _minErrorVersion = Number.MIN_SAFE_INTEGER;
  protected readonly _maxErrorVersion = Number.MAX_SAFE_INTEGER;

  protected getNodeNotSupportedHelpLink(): Promise<string> {
    return Promise.resolve(v3NodeNotSupportedHelpLink);
  }

  protected async getSupportedVersions(projectPath?: string): Promise<string[]> {
    if (!projectPath) {
      return [];
    }
    const supportedVersion = await this.getSupportedVersion(projectPath);
    return supportedVersion ? [supportedVersion] : [];
  }

  private async getSupportedVersion(projectPath: string): Promise<string | undefined> {
    try {
      const packageJson = await fs.readJSON(path.join(projectPath, "package.json"));
      const node = packageJson?.engines?.node;
      if (typeof node !== "string") {
        return undefined;
      }
      return node;
    } catch {
      return undefined;
    }
  }

  protected isVersionSupported(supportedVersions: string[], version: NodeVersion): boolean {
    if (supportedVersions.length == 0) {
      return true;
    }
    const supportedVersion = supportedVersions[0];
    return semver.satisfies(version.version, supportedVersion);
  }

  protected getVersionNotSupportedError(
    supportedVersions: string[],
    version: NodeVersion
  ): UserError {
    const supportedVersionsString = supportedVersions.join(", ");
    return new NodejsNotRecommendedError(version.version, supportedVersionsString);
  }
}

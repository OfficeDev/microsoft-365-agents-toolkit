// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AxiosInstance, AxiosResponse, default as axios } from "axios";
import * as path from "path";
import * as os from "os";
import {
  ConfigFolderName,
  FxError,
  LogProvider,
  SystemError,
  TelemetryReporter,
  UserError,
} from "@microsoft/teamsfx-api";
import * as fs from "fs-extra";
import { cpUtils } from "../../../utils/depsChecker/cpUtils";
import { finished, Readable, Writable } from "stream";
import {
  DepsCheckerEvent,
  Messages,
  TelemetryMeasurement,
  TelemtryMessages,
} from "../../../utils/depsChecker/common";
import {
  SolutionTelemetryComponentName,
  SolutionTelemetryProperty,
  SolutionTelemetrySuccess,
} from "../../../constants";

import { performance } from "perf_hooks";
import { DriverContext } from "../../interface/commonArgs";
import { InstallSoftwareError } from "../../../../error/common";
import { DownloadBicepCliError } from "../../../../error/arm";
import { isMacOS, isWindows } from "../../../deps-checker/util/system";
import { maskSecret } from "../../../../common/stringUtils";
import { TelemetryProperty } from "../../../../common/telemetry";

const BicepName = "Bicep";

const timeout = 5 * 60 * 1000;
const source = "bicep-envchecker";
const bicepReleaseApiUrl = "https://api.github.com/repos/Azure/bicep/releases";

export async function ensureBicepForDriver(ctx: DriverContext, version: string): Promise<string> {
  const bicepChecker = new BicepChecker(version, ctx.logProvider, ctx.telemetryReporter);
  try {
    const isPrivateBicepInstalled: boolean = await bicepChecker.isPrivateBicepInstalled();
    if (!isPrivateBicepInstalled) {
      await bicepChecker.install();
    }
  } catch (err) {
    throw err;
  }
  return bicepChecker.getBicepExecPath();
}

class BicepChecker {
  private readonly _logger: LogProvider | undefined;
  private readonly _telemetry: TelemetryReporter | undefined;
  private readonly version: string;
  private readonly _axios: AxiosInstance;

  constructor(version: string, logger?: LogProvider, telemetry?: TelemetryReporter) {
    this._logger = logger;
    this._telemetry = telemetry;
    this.version = version;
    this._axios = axios.create({
      headers: { "content-type": "application/json" },
      timeout: timeout,
      timeoutErrorMessage: "Failed to download bicep by http request timeout",
    });
  }

  public async install(): Promise<void> {
    await this.cleanup();

    await this.installBicep();

    if (!(await this.validate())) {
      await this.handleInstallFailed();
    }
    await this.handleInstallCompleted();
  }

  async getVersions(): Promise<string[]> {
    const response: AxiosResponse<Array<{ tag_name: string }>> = await this._axios.get(
      bicepReleaseApiUrl,
      {
        headers: { Accept: "application/vnd.github.v3+json" },
      }
    );
    const versions = response.data.map((item) => item.tag_name);
    return versions;
  }

  private async cleanup() {
    try {
      await fs.emptyDir(this.getBicepInstallDir());
    } catch (err: any) {
      this._logger?.debug(
        `Failed to clean up path: ${this.getBicepInstallDir()}, error: ${err.toString() as string}`
      );
    }
  }

  private async installBicep(): Promise<void> {
    try {
      const start = performance.now();
      await this.doInstallBicep();
      this._telemetry?.sendTelemetryEvent(
        DepsCheckerEvent.bicepInstallScriptCompleted,
        getCommonProps(),
        {
          [TelemetryMeasurement.completionTime]: Number(
            ((performance.now() - start) / 1000).toFixed(2)
          ),
        }
      );
    } catch (err: any) {
      sendSystemErrorEvent(
        DepsCheckerEvent.bicepInstallScriptError,
        TelemtryMessages.failedToInstallBicep,
        err,
        this._telemetry
      );
      throw new DownloadBicepCliError(bicepReleaseApiUrl, err as Error);
    }
  }

  getBicepDisplayBicepName() {
    return `${BicepName} (${this.version})`;
  }

  private async doInstallBicep(): Promise<void> {
    const installDir = this.getBicepExecPath();

    this._logger?.info(
      Messages.downloadBicep()
        .replace("@NameVersion", `Bicep ${this.version}`)
        .replace("@InstallDir", installDir)
    );
    const axiosResponse = await this._axios.get(
      `https://github.com/Azure/bicep/releases/download/${
        this.version
      }/${this.getBicepBitSuffixName()}`,
      {
        responseType: "stream",
      }
    );

    const bicepReader: Readable = axiosResponse.data;
    const bicepWriter: Writable = fs.createWriteStream(installDir);
    // https://nodejs.org/api/fs.html#fscreatewritestreampath-options
    // on 'error' or 'finish' the file descriptor will be closed automatically
    // calling writer.end() again will hang
    await this.writeBicepBits(bicepWriter, bicepReader);
    fs.chmodSync(installDir, 0o755);
  }

  private async writeBicepBits(writer: Writable, reader: Readable): Promise<void> {
    return new Promise((resolve: (value: void) => void, reject: (e: Error) => void): void => {
      reader.on("error", (err) => {
        // Handles reader error.
        writer.end();
        reject(err);
      });

      // https://nodejs.org/api/stream.html#readablepipedestination-options
      // If the Readable stream emits an error during processing, the Writable destination is **NOT** closed.
      reader.pipe(writer);
      finished(writer, (err?: NodeJS.ErrnoException | null) => {
        // Handles writer end and writer error.
        // By handling writer end, it implicitly handles reader end because of reader.pipe(writer).
        // But reader error is not handled here.
        if (err) reject(err);
        else resolve();
      });
    });
  }

  private async validate(): Promise<boolean> {
    let isVersionSupported = false;
    let privateVersion = "";
    try {
      privateVersion = await this.queryVersion(this.getBicepExecPath());
      isVersionSupported = this.isVersionSupported(privateVersion);
    } catch (err: any) {
      sendSystemErrorEvent(
        DepsCheckerEvent.bicepValidationError,
        TelemtryMessages.failedToValidateBicep,
        err,
        this._telemetry
      );
      this._logger?.error(
        `${TelemtryMessages.failedToValidateBicep}, error = ${err.toString() as string}`
      );
    }

    if (!isVersionSupported) {
      this._telemetry?.sendTelemetryEvent(DepsCheckerEvent.bicepValidationError, {
        ...{ "bicep-private-version": privateVersion },
        ...getCommonProps(),
      });
    }

    return isVersionSupported;
  }

  private async handleInstallCompleted() {
    this._telemetry?.sendTelemetryEvent(DepsCheckerEvent.bicepInstallCompleted);
    this._logger?.info(
      Messages.finishInstallBicep().replace("@NameVersion", this.getBicepDisplayBicepName())
    );
    return Promise.resolve();
  }

  private async handleInstallFailed(): Promise<void> {
    await this.cleanup();
    this._telemetry?.sendTelemetryErrorEvent(DepsCheckerEvent.bicepInstallError);
    throw new InstallSoftwareError(source, this.getBicepDisplayBicepName());
  }

  private isVersionSupported(version: string): boolean {
    return this.version === version;
  }

  public async isPrivateBicepInstalled(): Promise<boolean> {
    try {
      const version = await this.queryVersion(this.getBicepExecPath());
      return this.isVersionSupported(version);
    } catch (e) {
      // do nothing
      return false;
    }
  }

  public getBicepExecPath(): string {
    return path.join(this.getBicepInstallDir(), this.getBicepFileName());
  }

  private getBicepFileName(): string {
    if (isWindows()) {
      return "bicep.exe";
    }
    return "bicep";
  }

  private getBicepBitSuffixName(): string {
    if (isWindows()) {
      return "bicep-win-x64.exe";
    }
    if (isMacOS()) {
      return "bicep-osx-x64";
    }
    if (fs.pathExistsSync("/lib/ld-musl-x86_64.so.1")) {
      return "bicep-linux-musl-x64";
    }
    return "bicep-linux-x64";
  }

  private getBicepInstallDir(): string {
    return path.join(os.homedir(), `.${ConfigFolderName}`, "bin", "bicep", this.version);
  }

  private async queryVersion(path: string): Promise<string> {
    const output = await cpUtils.executeCommand(
      undefined,
      this._logger,
      { shell: false },
      path,
      "--version"
    );
    const regex = /(?<major_version>\d+)\.(?<minor_version>\d+)\.(?<patch_version>\d+)/gm;
    const match = regex.exec(output);
    if (!match) {
      return "";
    }
    return `v${match.groups?.major_version || ""}.${match.groups?.minor_version || ""}.${
      match.groups?.patch_version || ""
    }`;
  }
}

function sendSystemErrorEvent(
  eventName: DepsCheckerEvent,
  errorMessage: string,
  errorStack: string,
  telemetry?: TelemetryReporter
): void {
  const error = new SystemError(
    source,
    eventName,
    `errorMsg=${errorMessage},errorStack=${errorStack}`
  );
  error.stack = errorStack;
  sendErrorTelemetryThenReturnError(eventName, error, telemetry, getCommonProps());
}

function getCommonProps(): { [key: string]: string } {
  const properties: { [key: string]: string } = {};
  properties[TelemetryMeasurement.OSArch] = os.arch();
  properties[TelemetryMeasurement.OSRelease] = os.release();
  properties[TelemetryProperty.Component] = SolutionTelemetryComponentName;
  properties[TelemetryProperty.Success] = SolutionTelemetrySuccess.Yes;
  return properties;
}

function sendErrorTelemetryThenReturnError(
  eventName: string,
  error: FxError,
  reporter?: TelemetryReporter,
  properties?: { [p: string]: string },
  measurements?: { [p: string]: number },
  errorProps?: string[]
): FxError {
  if (!properties) {
    properties = {};
  }

  if (TelemetryProperty.Component in properties === false) {
    properties[TelemetryProperty.Component] = SolutionTelemetryComponentName;
  }

  properties[TelemetryProperty.Success] = "no";
  if (error instanceof UserError) {
    properties[TelemetryProperty.ErrorType] = "user";
  } else {
    properties[TelemetryProperty.ErrorType] = "system";
  }

  properties[TelemetryProperty.ErrorCode] = `${error.source}.${error.name}`;
  properties[TelemetryProperty.ErrorMessage] = maskSecret(error.message);

  reporter?.sendTelemetryErrorEvent(eventName, properties, measurements, errorProps);
  return error;
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Xiaofu Huang <xiaofhua@microsoft.com>
 */
import { hooks } from "@feathersjs/hooks/lib";
import { FxError, Result } from "@microsoft/teamsfx-api";
import * as path from "path";
import semver from "semver";
import { Service } from "typedi";
import { ErrorContextMW } from "../../../common/globalVars";
import { getLocalizedString } from "../../../common/localizeUtils";
import { InvalidActionInputError } from "../../../error/common";
import {
  DependencyStatus,
  EmptyLogger,
  EmptyTelemetry,
  TestToolReleaseType,
  v3DefaultHelpLink,
} from "../../deps-checker";
import { DotnetChecker } from "../../deps-checker/internal/dotnetChecker";
import { FuncToolChecker } from "../../deps-checker/internal/funcToolChecker";
import { TestToolChecker } from "../../deps-checker/internal/testToolChecker";
import { LocalCertificate, LocalCertificateManager } from "../../local/localCertificateManager";
import { wrapRun } from "../../utils/common";
import { DriverContext } from "../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { WrapDriverContext } from "../util/wrapUtil";
import {
  Summaries,
  TelemetryDepsCheckStatus,
  TelemetryDevCertStatus,
  TelemetryProperties,
  toolsInstallDescription,
} from "./constant";
import { DotnetInstallationUserError } from "./error/dotnetInstallationUserError";
import { FuncInstallationUserError } from "./error/funcInstallationUserError";
import { TestToolInstallationUserError } from "./error/testToolInstallationUserError";
import { InstallToolArgs } from "./interfaces/InstallToolArgs";

const ACTION_NAME = "devTool/install";
const helpLink = "https://aka.ms/teamsfx-actions/devtool-install";

const outputKeys = {
  sslCertFile: "sslCertFile",
  sslKeyFile: "sslKeyFile",
  funcPath: "funcPath",
  dotnetPath: "dotnetPath",
  testToolPath: "testToolPath",
};

@Service(ACTION_NAME)
export class ToolsInstallDriver implements StepDriver {
  description = toolsInstallDescription();
  readonly progressTitle = getLocalizedString("driver.prerequisite.progressBar");

  async run(
    args: InstallToolArgs,
    context: DriverContext
  ): Promise<Result<Map<string, string>, FxError>> {
    const wrapContext = new WrapDriverContext(context, ACTION_NAME, ACTION_NAME);
    return await this._run(args, wrapContext);
  }

  async execute(
    args: InstallToolArgs,
    context: DriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<ExecutionResult> {
    const wrapContext = new WrapDriverContext(context, ACTION_NAME, ACTION_NAME);
    const result = await this._run(args, wrapContext, outputEnvVarNames);
    return {
      result: result,
      summaries: wrapContext.summaries,
    };
  }

  @hooks([addStartAndEndTelemetry(ACTION_NAME, ACTION_NAME)])
  async _run(
    args: InstallToolArgs,
    wrapContext: WrapDriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<Result<Map<string, string>, FxError>> {
    return wrapRun(async () => {
      const impl = new ToolsInstallDriverImpl(wrapContext);
      return await impl.run(args, outputEnvVarNames);
    }, ACTION_NAME);
  }
}

export class ToolsInstallDriverImpl {
  constructor(private context: WrapDriverContext) {}

  @hooks([ErrorContextMW({ source: "DevTools", component: "ToolsInstallDriverImpl" })])
  async run(
    args: InstallToolArgs,
    outputEnvVarNames?: Map<string, string>
  ): Promise<Map<string, string>> {
    const res = new Map<string, string>();
    this.validateArgs(args);

    this.setArgTelemetry(args);

    if (args.devCert) {
      const localCertRes = await this.resolveLocalCertificate(
        args.devCert.trust,
        outputEnvVarNames
      );
      localCertRes.forEach((v, k) => res.set(k, v));
    }

    if (args.func) {
      const funcRes = await this.resolveFuncCoreTools(
        `${args.func.version}`,
        args.func.symlinkDir,
        outputEnvVarNames
      );
      funcRes.forEach((v, k) => res.set(k, v));
    }

    if (args.dotnet) {
      const dotnetRes = await this.resolveDotnet(outputEnvVarNames);
      dotnetRes.forEach((v, k) => res.set(k, v));
    }

    if (args.testTool) {
      await this.resolveTestTool(
        // Hardcode to npm release type if running from YAML
        TestToolReleaseType.Npm,
        `${args.testTool.version}`,
        args.testTool.symlinkDir
      );
    }

    return res;
  }

  async resolveLocalCertificate(
    trustDevCert: boolean,
    outputEnvVarNames?: Map<string, string>
  ): Promise<Map<string, string>> {
    const res = new Map<string, string>();
    // Do not print any log in LocalCertificateManager, use the error message returned instead.
    const certManager = new LocalCertificateManager(this.context.ui);
    const localCertResult = await certManager.setupCertificate(trustDevCert);
    this.context.logProvider.debug(
      `Dev cert result: ${JSON.stringify({
        cert: localCertResult.certPath,
        key: localCertResult.keyPath,
        alreadyTrusted: localCertResult.alreadyTrusted,
        isTrusted: localCertResult.isTrusted,
      })}`
    );
    if (trustDevCert) {
      let name = outputEnvVarNames?.get(outputKeys.sslCertFile);
      if (name) {
        res.set(name, localCertResult.certPath);
      }
      name = outputEnvVarNames?.get(outputKeys.sslKeyFile);
      if (name) {
        res.set(name, localCertResult.keyPath);
      }
    }
    this.setDevCertTelemetry(trustDevCert, localCertResult);

    if (typeof localCertResult.isTrusted === "undefined") {
      this.context.logProvider.warning(Summaries.devCertSkipped());
      this.context.addSummary(Summaries.devCertSkipped());
    } else if (localCertResult.isTrusted === false) {
      throw localCertResult.error;
    } else {
      this.context.addSummary(Summaries.devCertSuccess(trustDevCert));
    }

    return res;
  }

  async resolveFuncCoreTools(
    version: string,
    symlinkDir?: string,
    outputEnvVarNames?: Map<string, string>
  ): Promise<Map<string, string>> {
    const res = new Map<string, string>();
    const funcToolChecker = new FuncToolChecker();
    const funcStatus = await funcToolChecker.resolve({
      version: version,
      symlinkDir: symlinkDir,
      projectPath: this.context.projectPath,
    });
    this.context.logProvider.debug(
      `Func tool result: ${JSON.stringify({
        isInstalled: funcStatus.isInstalled,
        version: funcStatus.details.installVersion,
        bin: funcStatus.details.binFolders,
        supportedVersions: funcStatus.details.supportedVersions,
      })}`
    );

    this.setDepsCheckTelemetry(TelemetryProperties.funcStatus, funcStatus);

    if (!funcStatus.isInstalled && funcStatus.error) {
      throw new FuncInstallationUserError(ACTION_NAME, funcStatus.error, funcStatus.error.helpLink);
    } else if (funcStatus.error) {
      this.context.logProvider.warning(funcStatus.error.message);
      this.context.addSummary(
        Summaries.funcSuccess(funcStatus.details.binFolders) + funcStatus.error.message
      );
    } else {
      this.context.addSummary(Summaries.funcSuccess(funcStatus.details.binFolders));
    }

    if (funcStatus?.details?.binFolders !== undefined) {
      const funcBinFolder = funcStatus.details.binFolders.join(path.delimiter);
      const name = outputEnvVarNames?.get(outputKeys.funcPath);
      if (name) {
        res.set(name, funcBinFolder);
      }
    }
    return res;
  }

  async resolveDotnet(outputEnvVarNames?: Map<string, string>): Promise<Map<string, string>> {
    const res = new Map<string, string>();
    const dotnetChecker = new DotnetChecker(new EmptyLogger(), new EmptyTelemetry());
    const dotnetStatus = await dotnetChecker.resolve();
    this.context.logProvider.debug(
      `.NET result: ${JSON.stringify({
        isInstalled: dotnetStatus.isInstalled,
        version: dotnetStatus.details?.installVersion,
        bin: dotnetStatus.details?.binFolders,
        supportedVersions: dotnetStatus.details?.supportedVersions,
      })}`
    );

    this.setDepsCheckTelemetry(TelemetryProperties.dotnetStatus, dotnetStatus);

    if (!dotnetStatus.isInstalled && dotnetStatus.error) {
      throw new DotnetInstallationUserError(
        ACTION_NAME,
        dotnetStatus.error,
        dotnetStatus.error.helpLink
      );
    } else if (dotnetStatus.error) {
      this.context.logProvider.warning(dotnetStatus.error?.message);
      this.context.addSummary(dotnetStatus.error?.message);
    } else {
      this.context.addSummary(Summaries.dotnetSuccess(dotnetStatus?.details?.binFolders));
    }
    if (dotnetStatus?.details?.binFolders !== undefined) {
      const dotnetBinFolder = `${dotnetStatus.details.binFolders
        .map((f: string) => path.dirname(f))
        .join(path.delimiter)}`;
      const name = outputEnvVarNames?.get(outputKeys.dotnetPath);
      if (name) {
        res.set(name, dotnetBinFolder);
      }
    }
    return res;
  }

  async resolveTestTool(
    releaseType: TestToolReleaseType,
    versionRange: string,
    symlinkDir: string
  ): Promise<void> {
    const checker = new TestToolChecker();
    const projectPath = this.context.projectPath;
    const status = await checker.resolve({
      releaseType,
      versionRange,
      symlinkDir,
      projectPath,
    });
    this.context.logProvider.debug(
      `Teams App Test Tool result: ${JSON.stringify({
        isInstalled: status.isInstalled,
        version: status.details.installVersion,
        bin: status.details.binFolders,
        supportedVersions: status.details.supportedVersions,
      })}`
    );

    this.setDepsCheckTelemetry(TelemetryProperties.testToolStatus, status);

    if (!status.isInstalled) {
      throw new TestToolInstallationUserError(
        ACTION_NAME,
        status.error,
        status.error?.helpLink || v3DefaultHelpLink
      );
    } else {
      this.context.addSummary(Summaries.testToolSuccess(status.details.binFolders));
    }
  }

  private validateArgs(args: InstallToolArgs): void {
    if (!!args.devCert && typeof args.devCert?.trust !== "boolean") {
      throw new InvalidActionInputError(ACTION_NAME, ["devCert.trust"], helpLink);
    }
    if (typeof args.func !== "undefined") {
      if (typeof args.func !== "object") {
        throw new InvalidActionInputError(ACTION_NAME, ["func"], helpLink);
      }
      if (
        (typeof args.func.version !== "string" && typeof args.func.version !== "number") ||
        !semver.validRange(`${args.func?.version}`)
      ) {
        throw new InvalidActionInputError(ACTION_NAME, ["func.version"], helpLink);
      }
      if (typeof args.func.symlinkDir !== "string" && typeof args.func.symlinkDir !== "undefined") {
        throw new InvalidActionInputError(ACTION_NAME, ["func.symlinkDir"], helpLink);
      }
    }
    if (!!args.dotnet && typeof args.dotnet !== "boolean") {
      throw new InvalidActionInputError(ACTION_NAME, ["dotnet"], helpLink);
    }
    if (typeof args.testTool !== "undefined") {
      if (typeof args.testTool !== "object") {
        throw new InvalidActionInputError(ACTION_NAME, ["testTool"], helpLink);
      }
      if (
        typeof args.testTool.version !== "string" ||
        !semver.validRange(`${args.testTool?.version}`)
      ) {
        throw new InvalidActionInputError(ACTION_NAME, ["testTool.version"], helpLink);
      }
      if (typeof args.testTool.symlinkDir !== "string") {
        throw new InvalidActionInputError(ACTION_NAME, ["testTool.symlinkDir"], helpLink);
      }
    }
  }

  private setArgTelemetry(args: InstallToolArgs): void {
    this.context.addTelemetryProperties({
      [TelemetryProperties.driverArgs]: JSON.stringify({
        devCert: args.devCert,
        func: {
          version: args.func?.version,
          symlinkDir: args.func?.symlinkDir
            ? path.resolve(args.func.symlinkDir) === path.resolve("./devTools/func")
              ? "<default>"
              : "<unknown>"
            : "<undefined>",
        },
        dotnet: args.dotnet,
        testTool: {
          version: args.testTool?.version,
          symlinkDir: args.testTool?.symlinkDir
            ? path.resolve(args.testTool.symlinkDir) === path.resolve("./devTools/testTool")
              ? "<default>"
              : "<unknown>"
            : "<undefined>",
        },
      }),
    });
  }

  private setDevCertTelemetry(trustDevCert: boolean, localCertResult: LocalCertificate) {
    this.context.addTelemetryProperties({
      [TelemetryProperties.devCertStatus]: !trustDevCert
        ? TelemetryDevCertStatus.Disabled
        : localCertResult.alreadyTrusted
        ? TelemetryDevCertStatus.AlreadyTrusted
        : localCertResult.isTrusted
        ? TelemetryDevCertStatus.Trusted
        : TelemetryDevCertStatus.NotTrusted,
    });
  }

  private setDepsCheckTelemetry(
    propertyName: typeof TelemetryProperties[keyof typeof TelemetryProperties],
    depStatus: DependencyStatus
  ): void {
    this.context.addTelemetryProperties({
      [propertyName]: depStatus.isInstalled
        ? depStatus.error
          ? TelemetryDepsCheckStatus.warn
          : TelemetryDepsCheckStatus.success
        : TelemetryDepsCheckStatus.failed,
    });
    if (depStatus.telemetryProperties) {
      this.context.addTelemetryProperties(depStatus.telemetryProperties);
    }
  }
}

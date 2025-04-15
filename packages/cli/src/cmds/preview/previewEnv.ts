// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import { Colors, err, FxError, LogLevel, ok, Result, signedOut } from "@microsoft/teamsfx-api";
import {
  AppStudioScopes,
  assembleError,
  QuestionNames,
  environmentNameManager,
  envUtil,
  FxCore,
  getSideloadingStatus,
  HubTypes,
  isValidProjectV3,
  loadTeamsFxDevScript,
  TelemetryContext,
} from "@microsoft/teamsfx-core";
import fs from "fs-extra";
import * as path from "path";
import * as util from "util";
import activate from "../../activate";
import cliLogger from "../../commonlib/log";
import M365TokenInstance from "../../commonlib/m365Login";
import { cliSource } from "../../constants";
import cliTelemetry from "../../telemetry/cliTelemetry";
import { TelemetryEvent, TelemetryProperty } from "../../telemetry/cliTelemetryEvents";
import CLIUIInstance from "../../userInteraction";
import { getColorizedString, getSystemInputs } from "../../utils";
import * as commonUtils from "./commonUtils";
import * as constants from "./constants";
import * as errors from "./errors";
import { openHubWebClientNew, openTeamsDesktopClient } from "./launch";
import { localTelemetryReporter } from "./localTelemetryReporter";
import { ServiceLogWriter } from "./serviceLogWriter";
import { Task } from "./task";
enum Progress {
  M365Account = "Microsoft 365 Account",
}

const ProgressMessage: { [key: string]: string } = Object.freeze({
  [Progress.M365Account]: `Checking ${Progress.M365Account}`,
});

export default class PreviewEnv {
  public readonly description = "Preview the current application.";

  protected runningTasks: Task[] = [];

  private readonly telemetryProperties: { [key: string]: string } = {};
  private readonly telemetryMeasurements: { [key: string]: number } = {};

  public async runCommand(args: {
    [argName: string]: boolean | string | string[] | undefined;
  }): Promise<Result<null, FxError>> {
    if (args.folder === undefined || !isValidProjectV3(args.folder as string)) {
      return err(errors.WorkspaceNotSupported(args.folder as string));
    }
    const workspaceFolder = path.resolve(args.folder as string);
    const env = args.env as string;
    const manifestFilePath =
      (args["manifest-file-path"] as string) ??
      path.join(workspaceFolder, "appPackage", "manifest.json");
    const runCommand: string | undefined = args["run-command"] as string;
    const runningPattern = args["running-pattern"] as string;
    const openOnly = args["open-only"] as boolean;
    const m365Host = args["m365-host"] as HubTypes;
    const execPath: string = args["exec-path"] as string;
    const browser = args.browser as constants.Browser;
    const browserArguments = (args["browser-arg"] as string[]) ?? [];
    const desktop = args["desktop"] as boolean;

    cliTelemetry.withRootFolder(workspaceFolder);
    this.telemetryProperties[TelemetryProperty.PreviewType] =
      environmentNameManager.isRemoteEnvironment(env.toLowerCase())
        ? `remote-${env}`
        : env.toLowerCase();
    this.telemetryProperties[TelemetryProperty.PreviewHub] = m365Host;
    this.telemetryProperties[TelemetryProperty.PreviewBrowser] = browser;

    return await localTelemetryReporter.runWithTelemetryGeneric(
      TelemetryEvent.Preview,
      async () =>
        this.doPreview(
          workspaceFolder,
          env,
          manifestFilePath,
          runCommand,
          runningPattern,
          openOnly,
          m365Host,
          browser,
          browserArguments,
          execPath,
          desktop
        ),
      (result: Result<null, FxError>, ctx: TelemetryContext) => {
        // whether on success or failure, send this.telemetryProperties and this.telemetryMeasurements
        Object.assign(ctx.properties, this.telemetryProperties);
        Object.assign(ctx.measurements, this.telemetryMeasurements);
        return result.isErr() ? result.error : undefined;
      },
      this.telemetryProperties
    );
  }

  async doPreview(
    workspaceFolder: string,
    env: string,
    manifestFilePath: string,
    runCommand: string | undefined,
    runningPattern: string,
    openOnly: boolean,
    hub: HubTypes,
    browser: constants.Browser,
    browserArguments: string[],
    execPath: string,
    desktop: boolean
  ): Promise<Result<null, FxError>> {
    // 1. load envs
    const envRes = await envUtil.readEnv(workspaceFolder, env, false, false);
    if (envRes.isErr()) {
      return err(envRes.error);
    }

    // 2. check m365 account
    const accountInfoRes = await localTelemetryReporter.runWithTelemetry(
      TelemetryEvent.PreviewPrereqsCheckM365Account,
      () => this.checkM365Account(process.env.TEAMS_APP_TENANT_ID)
    );
    if (accountInfoRes.isErr()) {
      return err(accountInfoRes.error);
    }

    // 3. previewWithManifest
    const urlRes = await this.previewWithManifest(workspaceFolder, env, hub, manifestFilePath);
    if (urlRes.isErr()) {
      return err(urlRes.error);
    }

    // 4. detect project type and set run-command, running-pattern
    if (
      !openOnly &&
      runCommand === undefined &&
      env.toLowerCase() === environmentNameManager.getLocalEnvName()
    ) {
      const runCommandRes = await this.detectRunCommand(workspaceFolder);
      if (runCommandRes.isErr()) {
        return err(runCommandRes.error);
      }
      runCommand = runCommandRes.value.runCommand;
      cliLogger.necessaryLog(
        LogLevel.Info,
        getColorizedString([
          { content: constants.runCommand.detectRunCommand, color: Colors.WHITE },
          { content: runCommand, color: Colors.BRIGHT_MAGENTA },
        ])
      );
    }
    runCommand = runCommand === "" ? undefined : runCommand;
    const runningPatternRegex =
      runningPattern !== undefined
        ? runningPattern === ""
          ? new RegExp(".*", "i")
          : new RegExp(runningPattern, "i")
        : constants.defaultRunningPattern;

    try {
      // 5. run command as background task
      this.runningTasks = [];
      if (
        runCommand !== undefined &&
        env.toLowerCase() === environmentNameManager.getLocalEnvName()
      ) {
        const runTaskRes = await localTelemetryReporter.runWithTelemetry(
          TelemetryEvent.PreviewStartServices,
          () => this.runCommandAsTask(workspaceFolder, runCommand!, runningPatternRegex, execPath)
        );
        if (runTaskRes.isErr()) {
          throw runTaskRes.error;
        }
      }

      // 6. open hub web client or Teams desktop client
      if (desktop && hub == HubTypes.teams) {
        const launchRes = await this.launchDesktopClient(
          env,
          urlRes.value,
          browser,
          browserArguments
        );
        if (launchRes.isErr()) {
          throw launchRes.error;
        }
      } else {
        const launchRes = await this.launchBrowser(
          env,
          hub,
          urlRes.value,
          browser,
          browserArguments
        );
        if (launchRes.isErr()) {
          throw launchRes.error;
        }
      }
      if (runCommand !== undefined && env === environmentNameManager.getLocalEnvName()) {
        cliLogger.necessaryLog(LogLevel.Warning, constants.waitCtrlPlusC);
      }
    } catch (error: any) {
      await this.shutDown();
      return err(error);
    }

    return ok(null);
  }

  async checkM365Account(appTenantId?: string): Promise<
    Result<
      {
        tenantId?: string;
        loginHint?: string;
      },
      FxError
    >
  > {
    let result = true;
    let summaryMsg = `${Progress.M365Account}`;
    let error = undefined;
    const accountBar = CLIUIInstance.createProgressBar(Progress.M365Account, 1);
    await accountBar.start(ProgressMessage[Progress.M365Account]);
    await accountBar.next(ProgressMessage[Progress.M365Account]);
    let loginHint: string | undefined = undefined;
    let tenantId: string | undefined = undefined;
    try {
      let loginStatusRes = await M365TokenInstance.getStatus({ scopes: AppStudioScopes });
      let token = loginStatusRes.isOk() ? loginStatusRes.value.token : undefined;
      if (loginStatusRes.isOk() && loginStatusRes.value.status === signedOut) {
        const tokenRes = await M365TokenInstance.getAccessToken({
          scopes: AppStudioScopes,
          showDialog: true,
        });
        token = tokenRes.isOk() ? tokenRes.value : undefined;
        loginStatusRes = await M365TokenInstance.getStatus({ scopes: AppStudioScopes });
      }
      if (token === undefined) {
        result = false;
        summaryMsg = constants.doctorResult.NotSignIn;
      } else {
        const isSideloadingEnabled = await getSideloadingStatus(token);
        if (isSideloadingEnabled === false) {
          // sideloading disabled
          result = false;
          summaryMsg = constants.doctorResult.SideLoadingDisabled;
        }
      }
      const tokenObject = loginStatusRes.isOk() ? loginStatusRes.value.accountInfo : undefined;
      if (tokenObject && tokenObject.upn) {
        loginHint = tokenObject.upn as string;
      }
      if (tokenObject && tokenObject.tid) {
        tenantId = tokenObject.tid as string;
      }
    } catch (err: any) {
      result = false;
      error = assembleError(err, cliSource);
    }
    if (result && loginHint) {
      summaryMsg = constants.doctorResult.SignInSuccess.split("@account").join(`${loginHint}`);
    }
    await accountBar.end(result);
    cliLogger.necessaryLog(LogLevel.Info, summaryMsg, true);
    if (!result) {
      return error ? err(error) : err(errors.PrerequisitesValidationM365AccountError(summaryMsg));
    }
    if (
      tenantId !== undefined &&
      appTenantId !== undefined &&
      tenantId.toLowerCase() !== appTenantId.toLowerCase()
    ) {
      cliLogger.necessaryLog(LogLevel.Warning, constants.m365SwitchedMessage);
    }
    return ok({ tenantId: tenantId, loginHint: loginHint });
  }

  protected async previewWithManifest(
    projectPath: string,
    env: string,
    hub: HubTypes,
    manifestFilePath: string
  ): Promise<Result<string, FxError>> {
    const coreRes = await activate(projectPath, true);
    const core = (coreRes as any).value as FxCore;
    const inputs = getSystemInputs(projectPath, env);
    inputs[QuestionNames.M365Host] = hub;
    inputs[QuestionNames.TeamsAppManifestFilePath] = manifestFilePath;
    // inputs[QuestionNames.ConfirmManifest] = "manifest"; // skip confirmation // confirm is skipped in question model
    return await core.previewWithManifest(inputs);
  }

  protected async detectRunCommand(projectPath: string): Promise<
    Result<
      {
        runCommand: string;
      },
      FxError
    >
  > {
    let runCommand: string | undefined = undefined;
    const hasPackageJson = await fs.pathExists(path.join(projectPath, "package.json"));
    const csprojs = (await fs.readdir(projectPath)).filter(
      (f) => path.extname(f).toLowerCase() === ".csproj"
    );
    const hasCsproj = csprojs.length === 1;
    if (hasPackageJson && !hasCsproj) {
      // package.json should have "dev:teamsfx"
      const script = await loadTeamsFxDevScript(projectPath);
      runCommand = script !== undefined ? "npm run dev:teamsfx" : undefined;
    } else if (!hasPackageJson && hasCsproj) {
      const csprojContent = await fs.readFile(path.join(projectPath, csprojs[0]), "utf-8");
      const isFunc =
        /sdk\s*=\s*"\s*microsoft\.net\.sdk\s*"/i.test(csprojContent) &&
        /packagereference.*=\s*"\s*microsoft\.net\.sdk\.functions\s*"/i.test(csprojContent);
      runCommand = isFunc ? "func start" : "dotnet run";
    }
    if (runCommand === undefined) {
      return err(errors.CannotDetectRunCommand());
    }
    return ok({ runCommand: runCommand });
  }

  protected async runCommandAsTask(
    projectPath: string,
    runCommand: string,
    runningPatternRegex: RegExp,
    execPath: string
  ): Promise<Result<null, FxError>> {
    const taskName = "Run Command";
    const execPathStr = execPath
      .split(path.delimiter)
      .map((subPath) => (path.isAbsolute(subPath) ? subPath : path.resolve(projectPath, subPath)))
      .join(path.delimiter);
    const runningTask = new Task(taskName, true, runCommand, undefined, {
      shell: true,
      cwd: projectPath,
      env: {
        PATH: `${execPathStr}${path.delimiter}${process.env.PATH}`,
      },
    });
    this.runningTasks.push(runningTask);
    const bar = CLIUIInstance.createProgressBar(taskName, 1);
    const startCb = commonUtils.createTaskStartCb(bar, runCommand, this.telemetryProperties);
    const stopCb = commonUtils.createTaskStopCb(bar, this.telemetryProperties);
    const serviceLogWriter = new ServiceLogWriter();
    await serviceLogWriter.init();
    cliLogger.necessaryLog(
      LogLevel.Info,
      getColorizedString([
        { content: `${taskName}: ${constants.runCommand.showWorkingFolder}`, color: Colors.WHITE },
        { content: projectPath, color: Colors.BRIGHT_GREEN },
      ])
    );
    cliLogger.necessaryLog(
      LogLevel.Info,
      getColorizedString([
        { content: `${taskName}: ${constants.runCommand.showCommand}`, color: Colors.WHITE },
        { content: runCommand, color: Colors.BRIGHT_MAGENTA },
      ])
    );
    cliLogger.necessaryLog(
      LogLevel.Info,
      getColorizedString([
        { content: `${taskName}: ${constants.runCommand.showRunningPattern}`, color: Colors.WHITE },
        { content: runningPatternRegex.toString(), color: Colors.BRIGHT_MAGENTA },
      ])
    );
    const taskRes = await runningTask.waitFor(
      runningPatternRegex,
      startCb,
      stopCb,
      undefined,
      serviceLogWriter
    );
    return taskRes.isOk() ? ok(null) : err(taskRes.error);
  }

  protected async launchBrowser(
    env: string,
    hub: HubTypes,
    url: string,
    browser: constants.Browser,
    browserArgs: string[]
  ): Promise<Result<null, FxError>> {
    await openHubWebClientNew(hub, url, browser, browserArgs, this.telemetryProperties);

    cliLogger.necessaryLog(
      LogLevel.Warning,
      util.format(constants.manifestChangesHintMessage(), `--env ${env}`)
    );

    return ok(null);
  }

  protected async launchDesktopClient(
    env: string,
    url: string,
    browser: constants.Browser,
    browserArgs: string[]
  ): Promise<Result<null, FxError>> {
    const loginStatusRes = await M365TokenInstance.getStatus({ scopes: AppStudioScopes });
    let username = "";
    if (
      loginStatusRes.isOk() &&
      loginStatusRes?.value?.accountInfo &&
      loginStatusRes?.value?.accountInfo["unique_name"]
    ) {
      username = " (" + (loginStatusRes.value.accountInfo["unique_name"] as string) + ")";
    }
    await openTeamsDesktopClient(url, username, browser, browserArgs, this.telemetryProperties);

    cliLogger.necessaryLog(
      LogLevel.Warning,
      util.format(constants.manifestChangesHintMessage(), `--env ${env}`)
    );

    return ok(null);
  }

  private async shutDown() {
    for (const task of this.runningTasks) {
      await task.terminate();
    }
  }
}

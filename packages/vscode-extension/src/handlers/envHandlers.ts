// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as vscode from "vscode";
import path from "path";
import * as util from "util";
import fs from "fs-extra";
import {
  FxError,
  Result,
  SingleSelectConfig,
  Stage,
  SystemError,
  UserError,
  Void,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import {
  isValidProject,
  InvalidProjectError,
  environmentManager,
  pathUtils,
} from "@microsoft/teamsfx-core";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../telemetry/extTelemetryEvents";
import { getTriggerFromProperty } from "../utils/telemetryUtils";
import { runCommand } from "./sharedOpts";
import envTreeProviderInstance from "../treeview/environmentTreeViewProvider";
import { workspaceUri } from "../globalVariables";
import { VS_CODE_UI } from "../qm/vsc_ui";
import { showError } from "../error/common";
import { ExtensionSource, ExtensionErrors } from "../error/error";
import { localize } from "../utils/localizeUtils";

export const envHandlersDeps = {
  isValidProject: (workspacePath: string | undefined) => isValidProject(workspacePath),
  listAllEnvConfigs: (projectPath: string) => environmentManager.listAllEnvConfigs(projectPath),
  getEnvFolderPath: (projectPath: string) => pathUtils.getEnvFolderPath(projectPath),
  pathExists: (filePath: string) => fs.pathExists(filePath),
  openTextDocument: (filePath: string) => vscode.workspace.openTextDocument(filePath),
  showTextDocument: (document: vscode.TextDocument) => vscode.window.showTextDocument(document),
  selectOption: (config: SingleSelectConfig) => VS_CODE_UI.selectOption(config),
  runCommand: (stage: Stage, args?: any) => runCommand(stage, args),
  reloadEnvironments: () => envTreeProviderInstance.reloadEnvironments(),
  sendTelemetryEvent: (eventName: string, properties?: any) =>
    ExtTelemetry.sendTelemetryEvent(eventName as any, properties),
  sendTelemetryErrorEvent: (eventName: string, error?: any) =>
    ExtTelemetry.sendTelemetryErrorEvent(eventName as any, error),
};

export async function createNewEnvironment(args?: any[]): Promise<Result<undefined, FxError>> {
  envHandlersDeps.sendTelemetryEvent(
    TelemetryEvent.CreateNewEnvironmentStart,
    getTriggerFromProperty(args)
  );
  const result = await envHandlersDeps.runCommand(Stage.createEnv);
  if (!result.isErr()) {
    await envHandlersDeps.reloadEnvironments();
  }
  return result;
}

export async function refreshEnvironment(args?: any[]): Promise<Result<Void, FxError>> {
  return await envHandlersDeps.reloadEnvironments();
}

export async function openConfigStateFile(args: any[]): Promise<any> {
  let telemetryStartName = TelemetryEvent.OpenManifestConfigStateStart;
  let telemetryName = TelemetryEvent.OpenManifestConfigState;

  if (args && args.length > 0 && args[0].from === "aad") {
    telemetryStartName = TelemetryEvent.OpenAadConfigStateStart;
    telemetryName = TelemetryEvent.OpenAadConfigState;
  }

  envHandlersDeps.sendTelemetryEvent(telemetryStartName);
  const workspacePath = workspaceUri?.fsPath;
  if (!workspacePath) {
    const noOpenWorkspaceError = new UserError(
      ExtensionSource,
      ExtensionErrors.NoWorkspaceError,
      localize("teamstoolkit.handlers.noOpenWorkspace")
    );
    void showError(noOpenWorkspaceError);
    envHandlersDeps.sendTelemetryErrorEvent(telemetryName, noOpenWorkspaceError);
    return err(noOpenWorkspaceError);
  }

  if (!envHandlersDeps.isValidProject(workspacePath)) {
    const invalidProjectError = new UserError(
      ExtensionSource,
      ExtensionErrors.InvalidProject,
      localize("teamstoolkit.handlers.invalidProject")
    );
    void showError(invalidProjectError);
    envHandlersDeps.sendTelemetryErrorEvent(telemetryName, invalidProjectError);
    return err(invalidProjectError);
  }

  let sourcePath: string | undefined = undefined;
  let env: string | undefined = undefined;
  if (args && args.length > 0) {
    env = args[0].env;
    if (!env) {
      const envRes: Result<string | undefined, FxError> = await askTargetEnvironment();
      if (envRes.isErr()) {
        envHandlersDeps.sendTelemetryErrorEvent(telemetryName, envRes.error);
        return err(envRes.error);
      }
      env = envRes.value;
    }

    // Load env folder from yml
    const envFolder = await envHandlersDeps.getEnvFolderPath(workspacePath);
    if (envFolder.isOk() && envFolder.value) {
      sourcePath = path.resolve(`${envFolder.value}/.env.${env as string}`);
    } else if (envFolder.isErr()) {
      return err(envFolder.error);
    }
  } else {
    const invalidArgsError = new SystemError(
      ExtensionSource,
      ExtensionErrors.InvalidArgs,
      util.format(localize("teamstoolkit.handlers.invalidArgs"), args ? JSON.stringify(args) : args)
    );
    void showError(invalidArgsError);
    envHandlersDeps.sendTelemetryErrorEvent(telemetryName, invalidArgsError);
    return err(invalidArgsError);
  }

  if (sourcePath && !(await envHandlersDeps.pathExists(sourcePath))) {
    const noEnvError = new UserError(
      ExtensionSource,
      ExtensionErrors.EnvFileNotFoundError,
      util.format(localize("teamstoolkit.handlers.findEnvFailed"), env)
    );
    void showError(noEnvError);
    envHandlersDeps.sendTelemetryErrorEvent(telemetryName, noEnvError);
    return err(noEnvError);
  }

  void envHandlersDeps.openTextDocument(sourcePath as string).then((document) => {
    void envHandlersDeps.showTextDocument(document);
  });
  envHandlersDeps.sendTelemetryEvent(telemetryName, {
    [TelemetryProperty.Success]: TelemetrySuccess.Yes,
  });
}

/**
 * Ask user to select environment, local is included
 */
export async function askTargetEnvironment(): Promise<Result<string, FxError>> {
  const projectPath = workspaceUri?.fsPath;
  if (!envHandlersDeps.isValidProject(projectPath)) {
    return err(new InvalidProjectError(projectPath || ""));
  }
  const envProfilesResult = await envHandlersDeps.listAllEnvConfigs(projectPath!);
  if (envProfilesResult.isErr()) {
    return err(envProfilesResult.error);
  }
  const config: SingleSelectConfig = {
    name: "targetEnvName",
    title: "Select an environment",
    options: envProfilesResult.value,
  };
  const selectedEnv = await envHandlersDeps.selectOption(config);
  if (selectedEnv.isErr()) {
    return err(selectedEnv.error);
  } else {
    return ok(selectedEnv.value.result as string);
  }
}

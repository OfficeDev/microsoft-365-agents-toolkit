// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  CreateProjectResult,
  err,
  FxError,
  Inputs,
  ok,
  Result,
  Stage,
  UserError,
} from "@microsoft/teamsfx-api";
import {
  ApiPluginStartOptions,
  AppStudioScopes,
  assembleError,
  AuthSvcScopes,
  CapabilityOptions,
  featureFlagManager,
  FeatureFlags,
  isUserCancelError,
  isValidOfficeAddInProject,
  QuestionNames,
  teamsDevPortalClient,
} from "@microsoft/teamsfx-core";
import * as vscode from "vscode";
import M365TokenInstance from "../commonlib/m365Login";
import { VS_CODE_UI } from "../qm/vsc_ui";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetryTriggerFrom,
} from "../telemetry/extTelemetryEvents";
import envTreeProviderInstance from "../treeview/environmentTreeViewProvider";
import { localize } from "../utils/localizeUtils";
import { getSystemInputs } from "../utils/systemEnvUtils";
import { getTriggerFromProperty } from "../utils/telemetryUtils";
import { openFolder, openOfficeDevFolder } from "../utils/workspaceUtils";
import { invokeTeamsAgent } from "./copilotChatHandlers";
import { runCommand } from "./sharedOpts";
import { ExtensionSource } from "../error/error";
import VsCodeLogInstance from "../commonlib/log";
import * as versionUtil from "../utils/versionUtil";
import { KiotaExtensionId, KiotaMinVersion } from "../constants";
import * as stringUtil from "util";

export async function createNewProjectHandler(...args: any[]): Promise<Result<any, FxError>> {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.CreateProjectStart, getTriggerFromProperty(args));
  let inputs: Inputs | undefined;
  let stage = Stage.create;
  if (args?.length === 1) {
    if (!!args[0].teamsAppFromTdp) {
      inputs = getSystemInputs();
      inputs.teamsAppFromTdp = args[0].teamsAppFromTdp;
      stage = Stage.createTdp;
    }
  } else if (args?.length === 2) {
    // from copilot chat
    inputs = { ...getSystemInputs(), ...args[1] };
  }
  const result = await runCommand(stage, inputs);
  if (result.isErr()) {
    return err(result.error);
  }

  const res = result.value as CreateProjectResult;

  // For Kiota integration
  if (
    featureFlagManager.getBooleanValue(FeatureFlags.KiotaIntegration) &&
    res.projectPath === "" &&
    res.lastCommand
  ) {
    return handleTriggerKiotaCommand(args, res, TelemetryEvent.CreateProject);
  }

  if (res.shouldInvokeTeamsAgent) {
    await invokeTeamsAgent([TelemetryTriggerFrom.CreateAppQuestionFlow]);
    return result;
  }
  const projectPathUri = vscode.Uri.file(res.projectPath);
  const isOfficeAddin = isValidOfficeAddInProject(projectPathUri.fsPath);
  // If it is triggered in @office /create for code gen, then do no open the temp folder.
  if (isOfficeAddin && inputs?.agent === "office") {
    return result;
  }
  // show local debug button by default
  if (isOfficeAddin) {
    await openOfficeDevFolder(projectPathUri, true, res.warnings, args);
  } else {
    await openFolder(projectPathUri, true, res.warnings, args);
  }
  return result;
}

export async function provisionHandler(...args: unknown[]): Promise<Result<unknown, FxError>> {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.ProvisionStart, getTriggerFromProperty(args));
  const result = await runCommand(Stage.provision);
  if (result.isErr() && isUserCancelError(result.error)) {
    return result;
  } else {
    // refresh env tree except provision cancelled.
    await envTreeProviderInstance.reloadEnvironments();
    return result;
  }
}

export async function deployHandler(...args: unknown[]): Promise<Result<null, FxError>> {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.DeployStart, getTriggerFromProperty(args));
  return await runCommand(Stage.deploy);
}

export async function publishHandler(...args: unknown[]): Promise<Result<null, FxError>> {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.PublishStart, getTriggerFromProperty(args));
  return await runCommand(Stage.publish);
}

export async function addWebpartHandler(...args: unknown[]) {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.AddWebpartStart, getTriggerFromProperty(args));
  return await runCommand(Stage.addWebpart);
}

export async function addPluginHandler(...args: unknown[]) {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.AddPluginStart, getTriggerFromProperty(args));
  const result = await runCommand(Stage.addPlugin);
  if (result.isErr()) {
    return err(result.error);
  }

  const res = result.value;
  if (featureFlagManager.getBooleanValue(FeatureFlags.KiotaIntegration) && res.lastCommand) {
    return handleTriggerKiotaCommand(args, res, TelemetryEvent.AddPlugin, res.projectPath);
  } else {
    return result;
  }
}

export async function addKnowledgeHandler(...args: unknown[]) {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.AddKnowledgeStart, getTriggerFromProperty(args));
  const result = await runCommand(Stage.addKnowledge);
  if (result.isErr()) {
    return err(result.error);
  }
  return result;
}

/**
 * scaffold based on app id from Developer Portal
 */
export async function scaffoldFromDeveloperPortalHandler(
  ...args: any[]
): Promise<Result<null, FxError>> {
  if (!args || args.length < 1) {
    // should never happen
    return ok(null);
  }

  const appId = args[0];
  const properties: { [p: string]: string } = {
    teamsAppId: appId,
  };

  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.HandleUrlFromDeveloperProtalStart, properties);
  const loginHint = args.length < 2 ? undefined : args[1];
  const progressBar = VS_CODE_UI.createProgressBar(
    localize("teamstoolkit.devPortalIntegration.checkM365Account.progressTitle"),
    1
  );

  await progressBar.start();
  let token = undefined;
  try {
    const tokenRes = await M365TokenInstance.signInWhenInitiatedFromTdp(
      { scopes: AppStudioScopes },
      loginHint
    );
    if (tokenRes.isErr()) {
      if ((tokenRes.error as any).displayMessage) {
        void vscode.window.showErrorMessage((tokenRes.error as any).displayMessage);
      } else {
        void vscode.window.showErrorMessage(
          localize("teamstoolkit.devPortalIntegration.generalError.message")
        );
      }
      ExtTelemetry.sendTelemetryErrorEvent(
        TelemetryEvent.HandleUrlFromDeveloperProtal,
        tokenRes.error,
        properties
      );
      await progressBar.end(false);
      return err(tokenRes.error);
    }
    token = tokenRes.value;

    // set region
    const AuthSvcTokenRes = await M365TokenInstance.getAccessToken({ scopes: AuthSvcScopes });
    if (AuthSvcTokenRes.isOk()) {
      await teamsDevPortalClient.setRegionEndpointByToken(AuthSvcTokenRes.value);
    }

    await progressBar.end(true);
  } catch (e) {
    void vscode.window.showErrorMessage(
      localize("teamstoolkit.devPortalIntegration.generalError.message")
    );
    await progressBar.end(false);
    const error = assembleError(e);
    ExtTelemetry.sendTelemetryErrorEvent(
      TelemetryEvent.HandleUrlFromDeveloperProtal,
      error,
      properties
    );
    return err(error);
  }

  let appDefinition;
  try {
    appDefinition = await teamsDevPortalClient.getApp(token, appId);
  } catch (error: any) {
    ExtTelemetry.sendTelemetryErrorEvent(
      TelemetryEvent.HandleUrlFromDeveloperProtal,
      error,
      properties
    );
    void vscode.window.showErrorMessage(
      localize("teamstoolkit.devPortalIntegration.getTeamsAppError.message")
    );
    return err(error);
  }

  const res = await createNewProjectHandler({ teamsAppFromTdp: appDefinition });

  if (res.isErr()) {
    ExtTelemetry.sendTelemetryErrorEvent(
      TelemetryEvent.HandleUrlFromDeveloperProtal,
      res.error,
      properties
    );
    return err(res.error);
  }

  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.HandleUrlFromDeveloperProtal, properties);
  return ok(null);
}

export async function copilotPluginAddAPIHandler(args: any[]) {
  // Telemetries are handled in runCommand()
  const inputs = getSystemInputs();
  if (args && args.length > 0) {
    const filePath = args[0].fsPath as string;
    const isFromApiPlugin: boolean = args[0].isFromApiPlugin ?? false;
    if (!isFromApiPlugin) {
      // Codelens for API ME. Trigger from manifest.json
      inputs[QuestionNames.ManifestPath] = filePath;
    } else {
      inputs[QuestionNames.ApiPluginType] = ApiPluginStartOptions.apiSpec().id;
      inputs[QuestionNames.DestinationApiSpecFilePath] = filePath;
      inputs[QuestionNames.ManifestPath] = args[0].manifestPath;
    }
  }
  const result = await runCommand(Stage.copilotPluginAddAPI, inputs);
  return result;
}

export async function addAuthActionHandler(...args: unknown[]) {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.AddAuthActionStart, getTriggerFromProperty(args));
  const inputs = getSystemInputs();
  const result = await runCommand(Stage.addAuthAction, inputs);
  void vscode.window
    .showInformationMessage(
      localize("teamstoolkit.handeler.addAuthConfig.notification"),
      localize("teamstoolkit.handeler.addAuthConfig.notification.provision")
    )
    .then((selection) => {
      if (selection === "Provision") {
        ExtTelemetry.sendTelemetryEvent(TelemetryEvent.ProvisionFromAddAuthConfig);
        void runCommand(Stage.provision);
      }
    });
  return result;
}

function handleTriggerKiotaCommand(
  args: any[],
  result: any,
  event: string,
  projectPath?: string
): Result<any, FxError> {
  if (!validateKiotaInstallation()) {
    void vscode.window
      .showInformationMessage(
        stringUtil.format(localize("teamstoolkit.error.KiotaNotInstalled"), KiotaMinVersion),
        "Install Kiota",
        "Cancel"
      )
      .then((selection) => {
        if (selection === "Install Kiota") {
          // Open market place to install kiota
          ExtTelemetry.sendTelemetryEvent(TelemetryEvent.InstallKiota, {
            ...getTriggerFromProperty(args),
          });
          void vscode.commands.executeCommand("extension.open", "ms-graph.kiota");
        } else {
          return err(
            new UserError(
              ExtensionSource,
              "KiotaNotInstalled",
              stringUtil.format(localize("teamstoolkit.error.KiotaNotInstalled"), KiotaMinVersion)
            )
          );
        }
      });

    ExtTelemetry.sendTelemetryEvent(event, {
      [TelemetryProperty.KiotaInstalled]: "No",
      ...getTriggerFromProperty(args),
    });
    VsCodeLogInstance.error(
      stringUtil.format(localize("teamstoolkit.error.KiotaNotInstalled"), KiotaMinVersion)
    );
    return ok({ projectPath: "" });
  } else {
    void vscode.commands.executeCommand("kiota.openApiExplorer.searchOrOpenApiDescription", {
      kind: "Plugin",
      type: "ApiPlugin",
      source: "ttk",
      ttkContext: {
        lastCommand: result.lastCommand,
        manifestPath: result.manifestPath,
      },
      projectPath: projectPath,
    });
    ExtTelemetry.sendTelemetryEvent(event, {
      [TelemetryProperty.KiotaInstalled]: "Yes",
      ...getTriggerFromProperty(args),
    });
    return ok(result);
  }
}

function validateKiotaInstallation(): boolean {
  const installed = vscode.extensions.getExtension(KiotaExtensionId);
  if (!installed) {
    return false;
  }

  const kiotaVersion = installed.packageJSON.version;
  if (!kiotaVersion) {
    return false;
  }

  return versionUtil.compare(kiotaVersion, KiotaMinVersion) !== -1;
}

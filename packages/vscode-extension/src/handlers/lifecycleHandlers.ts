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
  ActionStartOptions,
  AppStudioScopes,
  assembleError,
  AuthSvcScopes,
  featureFlagManager,
  FeatureFlags,
  isSovereignHigh,
  isUserCancelError,
  isValidOfficeAddInProject,
  QuestionNames,
  teamsDevPortalClient,
} from "@microsoft/teamsfx-core";
import * as stringUtil from "util";
import * as vscode from "vscode";
import VsCodeLogInstance from "../commonlib/log";
import M365TokenInstance from "../commonlib/m365Login";
import { ExtensionSource } from "../error/error";
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
import * as versionUtil from "../utils/versionUtil";
import { openFolder, openOfficeDevFolder } from "../utils/workspaceUtils";
import { invokeTeamsAgent } from "./copilotChatHandlers";
import { runCommand } from "./sharedOpts";
import { tools } from "../globalVariables";

export const lifecycleHandlersOps = {
  sendTelemetryEvent: (eventName: string, properties?: any) =>
    ExtTelemetry.sendTelemetryEvent(eventName as any, properties),
  sendTelemetryErrorEvent: (eventName: string, error: FxError, properties?: any) =>
    ExtTelemetry.sendTelemetryErrorEvent(eventName as any, error, properties),
  getTriggerFromProperty: (args?: any[]) => getTriggerFromProperty(args),
  getSystemInputs: () => getSystemInputs(),
  runCommand: (stage: Stage, inputs?: Inputs) => runCommand(stage, inputs),
  invokeTeamsAgent: (args: any[]) => invokeTeamsAgent(args),
  uriFile: (path: string) => vscode.Uri.file(path),
  showErrorMessage: (message: string) => vscode.window.showErrorMessage(message),
  showInformationMessage: (message: string, ...items: string[]) =>
    vscode.window.showInformationMessage(message, ...items),
  isValidOfficeAddInProject: (filePath: string) => isValidOfficeAddInProject(filePath),
  openOfficeDevFolder: (uri: vscode.Uri, openBrowser: boolean, warnings?: any[], args?: any[]) =>
    openOfficeDevFolder(uri, openBrowser, warnings, args),
  openFolder: (uri: vscode.Uri, openBrowser: boolean, warnings?: any[], args?: any[]) =>
    openFolder(uri, openBrowser, warnings, args),
  reloadEnvironments: () => envTreeProviderInstance.reloadEnvironments(),
  localize: (key: string, ...args: any[]) => localize(key, ...args),
  createProgressBar: (title: string, totalSteps: number) =>
    VS_CODE_UI.createProgressBar(title, totalSteps),
  signInWhenInitiatedFromTdp: (options: any, loginHint?: string) =>
    M365TokenInstance.signInWhenInitiatedFromTdp(options, loginHint ?? ""),
  isSovereignHigh: () => isSovereignHigh(),
  getAccessToken: (options: any) => M365TokenInstance.getAccessToken(options),
  setRegionEndpointByToken: (token: string) => teamsDevPortalClient.setRegionEndpointByToken(token),
  getApp: (token: string, appId: string) => teamsDevPortalClient.getApp(token, appId),
  getBooleanValue: (flag: Parameters<typeof featureFlagManager.getBooleanValue>[0]) =>
    featureFlagManager.getBooleanValue(flag),
  getTools: () => tools,
};
const lifecycleHandlersDeps = lifecycleHandlersOps;

export async function createNewProjectHandler(...args: any[]): Promise<Result<any, FxError>> {
  lifecycleHandlersDeps.sendTelemetryEvent(
    TelemetryEvent.CreateProjectStart,
    lifecycleHandlersDeps.getTriggerFromProperty(args)
  );
  let inputs: Inputs | undefined;
  let stage = Stage.create;
  if (args?.length === 1) {
    if (!!args[0].teamsAppFromTdp) {
      inputs = lifecycleHandlersDeps.getSystemInputs();
      inputs.teamsAppFromTdp = args[0].teamsAppFromTdp;
      stage = Stage.createTdp;
    }
  } else if (args?.length === 2 && args[0] !== TelemetryTriggerFrom.TreeView) {
    // from copilot chat or createDeclarativeAgentWithApiSpec
    inputs = { ...lifecycleHandlersDeps.getSystemInputs(), ...args[1] };
  }
  const result = await lifecycleHandlersDeps.runCommand(stage, inputs);
  if (result.isErr()) {
    return err(result.error);
  }

  const res = result.value as CreateProjectResult;

  if (res.shouldInvokeTeamsAgent) {
    await lifecycleHandlersDeps.invokeTeamsAgent([TelemetryTriggerFrom.CreateAppQuestionFlow]);
    return result;
  }
  const projectPathUri = lifecycleHandlersDeps.uriFile(res.projectPath);
  const isOfficeAddin = lifecycleHandlersDeps.isValidOfficeAddInProject(projectPathUri.fsPath);
  // If it is triggered in @office /create for code gen, then do no open the temp folder.
  if (isOfficeAddin && inputs?.agent === "office") {
    return result;
  }
  // show local debug button by default
  if (isOfficeAddin) {
    await lifecycleHandlersDeps.openOfficeDevFolder(projectPathUri, true, res.warnings, args);
  } else {
    await lifecycleHandlersDeps.openFolder(projectPathUri, true, res.warnings, args);
  }
  return result;
}

export async function provisionHandler(...args: unknown[]): Promise<Result<unknown, FxError>> {
  lifecycleHandlersDeps.sendTelemetryEvent(
    TelemetryEvent.ProvisionStart,
    lifecycleHandlersDeps.getTriggerFromProperty(args)
  );
  const result = await lifecycleHandlersDeps.runCommand(Stage.provision);
  if (result.isErr() && isUserCancelError(result.error)) {
    return result;
  } else {
    // refresh env tree except provision cancelled
    await lifecycleHandlersDeps.reloadEnvironments();
    return result;
  }
}

export async function deployHandler(...args: unknown[]): Promise<Result<null, FxError>> {
  lifecycleHandlersDeps.sendTelemetryEvent(
    TelemetryEvent.DeployStart,
    lifecycleHandlersDeps.getTriggerFromProperty(args)
  );
  return await lifecycleHandlersDeps.runCommand(Stage.deploy);
}

export async function publishHandler(...args: unknown[]): Promise<Result<null, FxError>> {
  lifecycleHandlersDeps.sendTelemetryEvent(
    TelemetryEvent.PublishStart,
    lifecycleHandlersDeps.getTriggerFromProperty(args)
  );
  return await lifecycleHandlersDeps.runCommand(Stage.publish);
}

export async function shareHandler(...args: unknown[]): Promise<Result<null, FxError>> {
  lifecycleHandlersDeps.sendTelemetryEvent(
    TelemetryEvent.ShareStart,
    lifecycleHandlersDeps.getTriggerFromProperty(args)
  );
  return await lifecycleHandlersDeps.runCommand(Stage.share);
}

export async function shareRemoveHandler(...args: unknown[]): Promise<Result<null, FxError>> {
  lifecycleHandlersDeps.sendTelemetryEvent(
    TelemetryEvent.ShareRemoveStart,
    lifecycleHandlersDeps.getTriggerFromProperty(args)
  );
  return await lifecycleHandlersDeps.runCommand(Stage.shareRemove);
}

export async function addWebpartHandler(...args: unknown[]) {
  lifecycleHandlersDeps.sendTelemetryEvent(
    TelemetryEvent.AddWebpartStart,
    lifecycleHandlersDeps.getTriggerFromProperty(args)
  );
  return await lifecycleHandlersDeps.runCommand(Stage.addWebpart);
}

export async function regeneratePluginHandler(...args: unknown[]) {
  lifecycleHandlersDeps.sendTelemetryEvent(
    TelemetryEvent.RegenerateActionStart,
    lifecycleHandlersDeps.getTriggerFromProperty(args)
  );
  const result = await lifecycleHandlersDeps.runCommand(Stage.RegeneratePlugin);
  if (result.isErr()) {
    return err(result.error);
  }
  return result;
}

export async function addPluginHandler(...args: unknown[]) {
  lifecycleHandlersDeps.sendTelemetryEvent(TelemetryEvent.AddPluginStart, {
    ...lifecycleHandlersDeps.getTriggerFromProperty(args),
    [TelemetryProperty.KiotaNPMIntegrationEnabled]: lifecycleHandlersDeps
      .getBooleanValue(FeatureFlags.KiotaNPMIntegration)
      .toString(),
  });
  const result = await lifecycleHandlersDeps.runCommand(Stage.addPlugin);
  if (result.isErr()) {
    return err(result.error);
  }
  // For the MCP "Add Action" flow, addPlugin only writes the URL to
  // .vscode/mcp.json. Reuse the same UX as the "DA with MCP" scaffolding
  // flow: open mcp.json and show the "start MCP server / Fetch Action"
  // notification so the user can populate tools afterwards.
  if (result.value && (result.value as { kind?: string }).kind === "mcp") {
    const { openWorkspaceMCPConfigHandler } = await import("./readmeHandlers");
    await openWorkspaceMCPConfigHandler(TelemetryTriggerFrom.Auto);
  }
  return result;
}

export async function metaOSExtendToDAHandler(...args: unknown[]) {
  lifecycleHandlersDeps.sendTelemetryEvent(
    TelemetryEvent.MetaOSExtendToDAStart,
    lifecycleHandlersDeps.getTriggerFromProperty(args)
  );

  const result = await lifecycleHandlersDeps.runCommand(Stage.metaOSExtendToDA);
  if (result.isErr()) {
    return err(result.error);
  }

  const projectPathUri = lifecycleHandlersDeps.uriFile(result.value.projectPath);
  await lifecycleHandlersDeps.openFolder(projectPathUri, true, result.value.warnings);
  return result;
}

export async function addKnowledgeHandler(...args: unknown[]) {
  lifecycleHandlersDeps.sendTelemetryEvent(
    TelemetryEvent.AddKnowledgeStart,
    lifecycleHandlersDeps.getTriggerFromProperty(args)
  );
  const result = await lifecycleHandlersDeps.runCommand(Stage.addKnowledge);
  if (result.isErr()) {
    return err(result.error);
  }
  return result;
}

export async function addSkillHandler(...args: unknown[]) {
  lifecycleHandlersDeps.sendTelemetryEvent(
    TelemetryEvent.AddSkillStart,
    lifecycleHandlersDeps.getTriggerFromProperty(args)
  );
  const result = await lifecycleHandlersDeps.runCommand(Stage.addSkill);
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

  lifecycleHandlersDeps.sendTelemetryEvent(
    TelemetryEvent.HandleUrlFromDeveloperProtalStart,
    properties
  );
  const loginHint = args.length < 2 ? undefined : args[1];
  const progressBar = lifecycleHandlersDeps.createProgressBar(
    localize("teamstoolkit.devPortalIntegration.checkM365Account.progressTitle"),
    1
  );

  await progressBar.start();
  let token = undefined;
  try {
    const tokenRes = await lifecycleHandlersDeps.signInWhenInitiatedFromTdp(
      { scopes: AppStudioScopes() },
      loginHint
    );
    if (tokenRes.isErr()) {
      if ((tokenRes.error as any).displayMessage) {
        void lifecycleHandlersDeps.showErrorMessage((tokenRes.error as any).displayMessage);
      } else {
        void lifecycleHandlersDeps.showErrorMessage(
          lifecycleHandlersDeps.localize("teamstoolkit.devPortalIntegration.generalError.message")
        );
      }
      lifecycleHandlersDeps.sendTelemetryErrorEvent(
        TelemetryEvent.HandleUrlFromDeveloperProtal,
        tokenRes.error,
        properties
      );
      await progressBar.end(false);
      return err(tokenRes.error);
    }
    token = tokenRes.value;

    if (!lifecycleHandlersDeps.isSovereignHigh()) {
      // set region
      const AuthSvcTokenRes = await lifecycleHandlersDeps.getAccessToken({
        scopes: AuthSvcScopes(),
      });
      if (AuthSvcTokenRes.isOk()) {
        await lifecycleHandlersDeps.setRegionEndpointByToken(AuthSvcTokenRes.value);
      }
    }

    await progressBar.end(true);
  } catch (e) {
    void lifecycleHandlersDeps.showErrorMessage(
      lifecycleHandlersDeps.localize("teamstoolkit.devPortalIntegration.generalError.message")
    );
    await progressBar.end(false);
    const error = assembleError(e);
    lifecycleHandlersDeps.sendTelemetryErrorEvent(
      TelemetryEvent.HandleUrlFromDeveloperProtal,
      error,
      properties
    );
    return err(error);
  }

  let appDefinition;
  try {
    appDefinition = await lifecycleHandlersDeps.getApp(token, appId);
  } catch (error: any) {
    lifecycleHandlersDeps.sendTelemetryErrorEvent(
      TelemetryEvent.HandleUrlFromDeveloperProtal,
      error,
      properties
    );
    void lifecycleHandlersDeps.showErrorMessage(
      lifecycleHandlersDeps.localize("teamstoolkit.devPortalIntegration.getTeamsAppError.message")
    );
    return err(error);
  }

  const res = await createNewProjectHandler({ teamsAppFromTdp: appDefinition });

  if (res.isErr()) {
    lifecycleHandlersDeps.sendTelemetryErrorEvent(
      TelemetryEvent.HandleUrlFromDeveloperProtal,
      res.error,
      properties
    );
    return err(res.error);
  }

  lifecycleHandlersDeps.sendTelemetryEvent(TelemetryEvent.HandleUrlFromDeveloperProtal, properties);
  return ok(null);
}

export async function copilotPluginAddAPIHandler(args: any[]) {
  // Telemetries are handled in runCommand()
  const inputs = lifecycleHandlersDeps.getSystemInputs();
  if (args && args.length > 0) {
    const filePath = args[0].fsPath as string;
    const isFromApiPlugin: boolean = args[0].isFromApiPlugin ?? false;
    if (!isFromApiPlugin) {
      // Codelens for API ME. Trigger from manifest.json
      inputs[QuestionNames.ManifestPath] = filePath;
    } else {
      inputs[QuestionNames.ActionType] = ActionStartOptions.apiSpec().id;
      inputs[QuestionNames.DestinationApiSpecFilePath] = filePath;
      inputs[QuestionNames.ManifestPath] = args[0].manifestPath;
    }
  }
  const result = await lifecycleHandlersDeps.runCommand(Stage.copilotPluginAddAPI, inputs);
  return result;
}

export async function setSensitivityLabelHandler(args: any[]) {
  lifecycleHandlersDeps.sendTelemetryEvent(
    TelemetryEvent.SetSensitivityLabelStart,
    lifecycleHandlersDeps.getTriggerFromProperty(args)
  );
  const inputs = lifecycleHandlersDeps.getSystemInputs();
  inputs[QuestionNames.DeclarativeAgentManifestPath] = args?.[0]?.declarativeAgentManifestPath;
  inputs[QuestionNames.SensitivityLabel] = args?.[0]?.sensitivityLabel;
  const result = await lifecycleHandlersDeps.runCommand(Stage.setSensitivityLabel, inputs);
  if (result.isErr()) {
    lifecycleHandlersDeps.sendTelemetryErrorEvent(
      TelemetryEvent.SetSensitivityLabel,
      result.error,
      lifecycleHandlersDeps.getTriggerFromProperty(args)
    );
    return;
  }
  lifecycleHandlersDeps.sendTelemetryEvent(
    TelemetryEvent.SetSensitivityLabel,
    lifecycleHandlersDeps.getTriggerFromProperty(args)
  );
  return;
}

export async function m365PreAuthHandler(args: any[]) {
  lifecycleHandlersDeps.sendTelemetryEvent(
    TelemetryEvent.m365PreAuthStart,
    lifecycleHandlersDeps.getTriggerFromProperty(args)
  );
  const res = await lifecycleHandlersDeps
    .getTools()
    .tokenProvider?.m365TokenProvider?.getAccessToken({
      scopes: args[0].scopes,
    });
  if (res.isErr()) {
    lifecycleHandlersDeps.sendTelemetryErrorEvent(
      TelemetryEvent.m365PreAuth,
      res.error,
      lifecycleHandlersDeps.getTriggerFromProperty(args)
    );
    return;
  }
  lifecycleHandlersDeps.sendTelemetryEvent(
    TelemetryEvent.m365PreAuth,
    lifecycleHandlersDeps.getTriggerFromProperty(args)
  );
  return;
}

export async function addAuthActionHandler(...args: unknown[]) {
  lifecycleHandlersDeps.sendTelemetryEvent(
    TelemetryEvent.AddAuthActionStart,
    lifecycleHandlersDeps.getTriggerFromProperty(args)
  );
  const inputs = lifecycleHandlersDeps.getSystemInputs();
  const result = await lifecycleHandlersDeps.runCommand(Stage.addAuthAction, inputs);
  void lifecycleHandlersDeps
    .showInformationMessage(
      lifecycleHandlersDeps.localize("teamstoolkit.handeler.addAuthConfig.notification"),
      lifecycleHandlersDeps.localize("teamstoolkit.handeler.addAuthConfig.notification.provision")
    )
    .then((selection) => {
      if (selection === "Provision") {
        lifecycleHandlersDeps.sendTelemetryEvent(TelemetryEvent.ProvisionFromAddAuthConfig);
        void lifecycleHandlersDeps.runCommand(Stage.provision);
      }
    });
  return result;
}

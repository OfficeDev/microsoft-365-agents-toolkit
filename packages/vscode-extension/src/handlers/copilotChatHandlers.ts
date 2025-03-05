// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as util from "util";
import * as vscode from "vscode";

import { FxError, Result, SystemError, UserError, err, ok } from "@microsoft/teamsfx-api";
import { assembleError, globalStateGet, globalStateUpdate } from "@microsoft/teamsfx-core";
import VsCodeLogInstance from "../commonlib/log";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
  TelemetryTriggerFrom,
} from "../telemetry/extTelemetryEvents";
import { getTriggerFromProperty } from "../utils/telemetryUtils";
import { showOutputChannelHandler } from "./showOutputChannel";
import { localize } from "../utils/localizeUtils";
import { GlobalKey, InstallCopilotChatLink } from "../constants";
import { isVSCodeInsiderVersion } from "../utils/versionUtil";
import { VS_CODE_UI } from "../qm/vsc_ui";

const githubCopilotChatExtensionId = "github.copilot-chat";
const teamsAgentLink = "https://aka.ms/install-teamsapp";

enum errorNames {
  NoActiveTextEditor = "NoActiveTextEditor",
  openCopilotError = "openCopilotError",
}

function githubCopilotInstalled(): boolean {
  const extension = vscode.extensions.getExtension(githubCopilotChatExtensionId);
  return !!extension;
}

async function openOutputInEditor(): Promise<void> {
  // try open output in editor but ignore error
  try {
    showOutputChannelHandler();
    await vscode.commands.executeCommand("workbench.action.openActiveLogOutputFile");
  } catch (e) {}
}

export async function openGithubCopilotChat(args?: any[]): Promise<Result<null, FxError>> {
  const startEventName = TelemetryEvent.OpenGitHubCopilotChatStart;
  const eventName = TelemetryEvent.openGitHubCopilotChat;
  const triggerFrom = getTriggerFromProperty(args);
  const hasQuery = !!args && args.length == 2;
  const query = hasQuery ? args[1] : "";

  const telemtryProperties = {
    ...triggerFrom,
    [TelemetryProperty.HasQueryForCopilotChat]: hasQuery.toString(),
  };
  ExtTelemetry.sendTelemetryEvent(startEventName, triggerFrom);
  try {
    await vscode.commands.executeCommand("workbench.panel.chat.view.copilot.focus");
    if (query) {
      const options = {
        query,
        isPartialQuery: true,
      };
      await vscode.commands.executeCommand("workbench.action.chat.open", options);
    } else {
      await vscode.commands.executeCommand("workbench.action.chat.open");
    }
    ExtTelemetry.sendTelemetryEvent(eventName, telemtryProperties);
    return ok(null);
  } catch (e) {
    const error = new SystemError(
      eventName,
      errorNames.openCopilotError,
      util.format(localize("teamstoolkit.handlers.chatTeamsAgentError", query)),
      util.format(localize("teamstoolkit.handlers.chatTeamsAgentError", query))
    );
    VsCodeLogInstance.error(error.message);
    ExtTelemetry.sendTelemetryErrorEvent(eventName, error, telemtryProperties);

    const assembledError = assembleError(e);
    if (assembledError.message) {
      VsCodeLogInstance.error(assembledError.message);
    }

    return err(error);
  }
}

export async function installGithubCopilotChatExtension(
  args?: any[]
): Promise<Result<null, FxError>> {
  const startEventName = TelemetryEvent.InstallCopilotChatStart;
  const eventName = TelemetryEvent.InstallCopilotChat;

  const isExtensionInstalled = githubCopilotInstalled();
  if (isExtensionInstalled) {
    void vscode.window.showInformationMessage(
      localize("teamstoolkit.handlers.installCopilotChatExtensionAlreadyInstalled")
    );
    return ok(null);
  }
  const telemetryProperties = getTriggerFromProperty(args);
  ExtTelemetry.sendTelemetryEvent(startEventName, telemetryProperties);
  try {
    await vscode.commands.executeCommand(
      "workbench.extensions.installExtension",
      githubCopilotChatExtensionId,
      {
        installPreReleaseVersion: isVSCodeInsiderVersion(), // VSCode insider need to install Github Copilot Chat of pre-release version
        enable: true,
      }
    );

    ExtTelemetry.sendTelemetryEvent(eventName, {
      ...telemetryProperties,
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
    });

    return ok(null);
  } catch (e) {
    const error = new SystemError(
      eventName,
      "InstallCopilotError",
      util.format(localize("teamstoolkit.handlers.installCopilotError", InstallCopilotChatLink)),
      util.format(localize("teamstoolkit.handlers.installCopilotError", InstallCopilotChatLink))
    );
    VsCodeLogInstance.error(error.message);
    ExtTelemetry.sendTelemetryErrorEvent(eventName, error, telemetryProperties);

    const assembledError = assembleError(e);
    if (assembledError.message) {
      VsCodeLogInstance.error(assembledError.message);
    }

    return err(error);
  }
}

export async function openInstallTeamsAgent(args?: any[]) {
  const startEventName = TelemetryEvent.OpenInstallTeamsAgentStart;
  const eventName = TelemetryEvent.OpenInstallTeamsAgent;

  const telemetryProperties = getTriggerFromProperty(args);
  ExtTelemetry.sendTelemetryEvent(startEventName, telemetryProperties);
  const openUrlRes = await VS_CODE_UI.openUrl(teamsAgentLink);
  if (openUrlRes.isOk()) {
    ExtTelemetry.sendTelemetryEvent(eventName, telemetryProperties);
  } else {
    ExtTelemetry.sendTelemetryErrorEvent(eventName, openUrlRes.error, telemetryProperties);
    VsCodeLogInstance.error(openUrlRes.error.message);
  }
}

export async function markTeamsAgentInstallationDone(args?: any[]) {
  const startEventName = TelemetryEvent.MarkTeamsAgentInstallationDoneStart;
  const eventName = TelemetryEvent.MarkTeamsAgentInstallationDone;

  ExtTelemetry.sendTelemetryEvent(startEventName);

  try {
    await globalStateUpdate(GlobalKey.TeamsAgentInstalled, true);
    ExtTelemetry.sendTelemetryEvent(eventName);
  } catch (e) {
    ExtTelemetry.sendTelemetryErrorEvent(eventName, assembleError(e));
  }
}

export async function openTeamsAgentWalkthrough(args?: any[]) {
  const triggerFromProperty = getTriggerFromProperty(args);
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.OpenTeamsAgentWalkthrough, triggerFromProperty);
  await vscode.commands.executeCommand("workbench.action.openWalkthrough", {
    category: "TeamsDevApp.ms-teams-vscode-extension#teamsAgentGetStarted",
  });
}

/**
 * Invoke @teamsapp
 * @param query query
 * @param triggerFromProperty trigger-from property
 * @param skipPreCheck skip pre-check or not. Default value is false.
 * @returns A boolean value indicates whether the query is sent or not. If not, it means the walkthrough is opened instead.
 */
async function invoke(
  query: string,
  triggerFromProperty: { [key: string]: TelemetryTriggerFrom },
  skipPreCheck = false
): Promise<Result<boolean, FxError>> {
  if (skipPreCheck) {
    const res = await openGithubCopilotChat([
      triggerFromProperty[TelemetryProperty.TriggerFrom],
      query,
    ]);
    if (res.isErr()) {
      return err(res.error);
    } else {
      return ok(true);
    }
  }

  const hasGitHubCopilotInstalled = githubCopilotInstalled();
  const hasTeamsAgentInstalled = await globalStateGet(GlobalKey.TeamsAgentInstalled, false);
  const hasGitHubCopilotSetup = await isGithubLoggedIn();

  if (hasGitHubCopilotInstalled && hasTeamsAgentInstalled && hasGitHubCopilotSetup) {
    if (triggerFromProperty[TelemetryProperty.TriggerFrom] === TelemetryTriggerFrom.Notification) {
      await openOutputInEditor();
    }

    const res = await openGithubCopilotChat([
      triggerFromProperty[TelemetryProperty.TriggerFrom],
      query,
    ]);
    if (res.isErr()) {
      return err(res.error);
    } else {
      return ok(true);
    }
  } else {
    await openTeamsAgentWalkthrough([triggerFromProperty[TelemetryProperty.TriggerFrom]]);
    return ok(false);
  }
}

/**
 * Invokes GitHub Copilot Chat for creating new app or development questions.
 * @param args args
 * @returns Result
 */
export async function invokeTeamsAgent(args?: any[]): Promise<Result<boolean, FxError>> {
  const eventName = TelemetryEvent.InvokeTeamsAgent;
  const triggerFromProperty = getTriggerFromProperty(args);
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.InvokeTeamsAgentStart, triggerFromProperty);

  let query = "";
  let shouldSkipPreCheck = false;
  switch (triggerFromProperty[TelemetryProperty.TriggerFrom]) {
    case TelemetryTriggerFrom.TreeView:
    case TelemetryTriggerFrom.CommandPalette:
      query =
        "@teamsapp Use this GitHub Copilot extension to ask questions about Teams app and agent development.";
      break;
    case TelemetryTriggerFrom.TeamsAgentWalkthroughExplore:
      shouldSkipPreCheck = true;
      query = "@teamsapp What's the difference between declarative and custom agents?";
      break;
    case TelemetryTriggerFrom.TeamsAgentWalkthroughCreate:
      shouldSkipPreCheck = true;
      query = "@teamsapp I want to create a ToDo Teams app.";
      break;
    case TelemetryTriggerFrom.TeamsAgentWalkthroughTroubleshoot:
      shouldSkipPreCheck = true;
      query = "@teamsapp My Teams app doesn't sideload when debugging with Teams Toolkit.";
      break;
    case TelemetryTriggerFrom.WalkThrough:
      shouldSkipPreCheck = true;
      query = "@teamsapp What can you do?";
      break;
    default:
      query =
        "@teamsapp Write your own query message to find relevant templates or samples to build your Teams app and agent as per your description. E.g. @teamsapp create an AI assistant bot that can complete common tasks.";
  }

  const res = await invoke(query, triggerFromProperty, shouldSkipPreCheck);

  if (res.isErr()) {
    ExtTelemetry.sendTelemetryErrorEvent(eventName, res.error, triggerFromProperty);
  } else {
    ExtTelemetry.sendTelemetryEvent(eventName, {
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
      ...triggerFromProperty,
      [TelemetryProperty.CopilotChatQuerySent]: res.value.toString(),
    });
  }
  return res;
}

async function isGithubLoggedIn(): Promise<boolean> {
  try {
    const accounts = await vscode.authentication.getAccounts("github");
    return accounts.length > 0;
  } catch (e) {
    return true;
  }
}

/**
 * Invokes teams agent for troubleshooting based on selected text.
 * @param args
 * @returns Result
 */
export async function troubleshootSelectedText(args?: any[]): Promise<Result<boolean, FxError>> {
  const eventName = TelemetryEvent.TroubleshootSelectedText;
  const triggerFromProperty = getTriggerFromProperty([TelemetryTriggerFrom.EditorContextMenu]);
  ExtTelemetry.sendTelemetryEvent(
    TelemetryEvent.TroubleshootSelectedTextStart,
    triggerFromProperty
  );

  const editor = vscode.window.activeTextEditor;
  let selectedText = "";
  if (editor) {
    const selection = editor.selection;
    selectedText = editor.document.getText(selection);
  } else {
    return err(
      new UserError(
        eventName,
        errorNames.NoActiveTextEditor,
        localize("teamstoolkit.handlers.teamsAgentTroubleshoot.noActiveEditor")
      )
    );
  }

  const query = `@teamsapp I'm encountering an issue in Teams Toolkit.
\`\`\`
{
  Error context: ${selectedText}
}
\`\`\`
Can you help me diagnose the issue and suggest possible solutions?
`;
  const res = await invoke(query, triggerFromProperty);

  if (res.isErr()) {
    ExtTelemetry.sendTelemetryErrorEvent(eventName, res.error, triggerFromProperty);
  } else {
    ExtTelemetry.sendTelemetryEvent(eventName, {
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
      ...triggerFromProperty,
      [TelemetryProperty.CopilotChatQuerySent]: res.value.toString(),
    });
  }
  return res;
}

/**
 * Invokes teams agent for troubleshooting current error.
 * @param args
 * @returns Result
 */
export async function troubleshootError(args?: any[]): Promise<Result<boolean, FxError>> {
  const eventName = TelemetryEvent.TroubleshootErrorFromNotification;
  if (!args || args.length !== 2) {
    // should never happen
    return ok(false);
  }

  const currentError = args[1] as FxError;
  const errorCode = `${currentError.source}.${currentError.name}`;
  const triggerFromProperty = getTriggerFromProperty(args);
  const telemtryProperties = {
    ...triggerFromProperty,
    [TelemetryProperty.ErrorCode]: errorCode,
  };
  ExtTelemetry.sendTelemetryEvent(
    TelemetryEvent.TroubleshootErrorFromNotificationStart,
    telemtryProperties
  );

  const query = `@teamsapp I'm encountering the following error in Teams Toolkit.
  \`\`\`
  {
    Error code: ${errorCode}
    Error message: ${currentError.message}
  }
  \`\`\`
  Can you help me diagnose the issue and suggest possible solutions?
  `;
  const res = await invoke(query, triggerFromProperty);

  if (res.isErr()) {
    ExtTelemetry.sendTelemetryErrorEvent(eventName, res.error, telemtryProperties);
  } else {
    ExtTelemetry.sendTelemetryEvent(eventName, {
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
      ...telemtryProperties,
      [TelemetryProperty.CopilotChatQuerySent]: res.value.toString(),
    });
  }
  return res;
}

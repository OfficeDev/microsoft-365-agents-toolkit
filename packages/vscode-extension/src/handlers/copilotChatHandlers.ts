// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as util from "util";
import * as vscode from "vscode";

import { FxError, Result, SystemError, UserError, err, ok } from "@microsoft/teamsfx-api";
import { assembleError, globalStateGet, globalStateUpdate } from "@microsoft/teamsfx-core";
import { UserCancelError, sleep } from "@microsoft/vscode-ui";
import VsCodeLogInstance from "../commonlib/log";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
  TelemetryTriggerFrom,
} from "../telemetry/extTelemetryEvents";
import { getTriggerFromProperty } from "../utils/telemetryUtils";
import { localize } from "../utils/localizeUtils";
import { showOutputChannelHandler } from "./showOutputChannel";
import { GlobalKey, InstallCopilotChatLink } from "../constants";
import { isVSCodeInsiderVersion } from "../utils/versionUtil";
import { VS_CODE_UI } from "../qm/vsc_ui";

const githubCopilotChatExtensionId = "github.copilot-chat";
const teamsAgentLink = "https://aka.ms/install-teamsapp";

enum errorNames {
  NoActiveTextEditor = "NoActiveTextEditor",
  CannotVerifyGithubCopilotChat = "CannotVerifyGithubCopilotChat",
  openCopilotError = "openCopilotError",
}

function githubCopilotInstalled(): boolean {
  const extension = vscode.extensions.getExtension(githubCopilotChatExtensionId);
  return !!extension;
}

async function openGithubCopilotChat(query: string): Promise<Result<null, FxError>> {
  const eventName = "openCopilotChat";
  try {
    const options = {
      query,
      isPartialQuery: true,
    };
    await vscode.commands.executeCommand("workbench.panel.chat.view.copilot.focus");
    await vscode.commands.executeCommand("workbench.action.chat.open", options);
    return ok(null);
  } catch (e) {
    const error = new SystemError(
      eventName,
      errorNames.openCopilotError,
      util.format(localize("teamstoolkit.handlers.chatTeamsAgentError", query)),
      util.format(localize("teamstoolkit.handlers.chatTeamsAgentError", query))
    );
    VsCodeLogInstance.error(error.message);
    ExtTelemetry.sendTelemetryErrorEvent(eventName, error);

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
  const startEventName = "install-copilot-chat-start";
  const eventName = "install-copilot-chat";

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
  const startEventName = "open-install-teams-agent-start";
  const eventName = "open-install-teams-agent";

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
  const startEventName = "mark-teams-agent-installation-done-start";
  const eventName = "mark-teams-agent-installation-done";

  const telemetryProperties = getTriggerFromProperty(args);
  ExtTelemetry.sendTelemetryEvent(startEventName, telemetryProperties);
  await globalStateUpdate(GlobalKey.TemasAgentInstalled, true);
  ExtTelemetry.sendTelemetryEvent(eventName, telemetryProperties);
}

export async function handleInstallTeamsAgentSelection(
  selection: string | undefined,
  telemetryProperties: {
    [key: string]: string;
  }
) {
  const eventName = "install-teams-agent-notification";
  const selectionTelemetryPropertyName = "selection";
  if (selection === localize("teamstoolkit.handlers.askInstallTeamsAgent.install")) {
    const installTelemetryProperties = {
      ...telemetryProperties,
      [selectionTelemetryPropertyName]: "install",
    };
    const openUrlRes = await VS_CODE_UI.openUrl(teamsAgentLink);
    if (openUrlRes.isOk()) {
      ExtTelemetry.sendTelemetryEvent(eventName, {
        ...installTelemetryProperties,
      });
    } else {
      ExtTelemetry.sendTelemetryErrorEvent(eventName, openUrlRes.error, installTelemetryProperties);
      VsCodeLogInstance.error(openUrlRes.error.message);
    }
  } else if (selection === localize("teamstoolkit.handlers.askInstallTeamsAgent.confirmInstall")) {
    ExtTelemetry.sendTelemetryEvent(eventName, {
      ...telemetryProperties,
      [selectionTelemetryPropertyName]: "confirmed",
    });
    await globalStateUpdate(GlobalKey.DoNotRemindInstallTeamsAgent, true);
  } else {
    const error = new UserCancelError(eventName, "cancel");
    ExtTelemetry.sendTelemetryErrorEvent(eventName, error, telemetryProperties);
  }
}

async function invoke(
  query: string,
  eventName: string,
  triggerFromProperty: { [key: string]: TelemetryTriggerFrom }
): Promise<Result<null, FxError>> {
  // const skipRemindInstallTeamsAgent = await globalStateGet(
  //   GlobalKey.DoNotRemindInstallTeamsAgent,
  //   false
  // );

  // console.log(triggerFromProperty[TelemetryProperty.TriggerFrom]);

  await vscode.commands.executeCommand(
    "setContext",
    "fx-extension.teamsAgentTriggerFrom",
    triggerFromProperty[TelemetryProperty.TriggerFrom]
  );

  // const welcomePageConfig = vscode.workspace.getConfiguration('workbench.welcomePage');
  // const configJSON = JSON.stringify(welcomePageConfig, null, 2); // Pretty-print with 2 spaces
  // console.log('Welcome Page Configuration (JSON):', configJSON);

  // const config = vscode.workspace.getConfiguration('workbench.welcomePage');
  // const walkthroughState = config.get<any>('walkthroughs');
  // if (walkthroughState && walkthroughState['TeamsDevApp.ms-teams-vscode-extension#buildIntelligentApps']) {
  //   vscode.window.showInformationMessage('Walkthrough is completed!');
  // } else {
  //   vscode.window.showInformationMessage('Walkthrough is not completed yet.');
  // }

  // await vscode.commands.executeCommand(
  //   "workbench.action.openWalkthrough",
  //   "TeamsDevApp.ms-teams-vscode-extension#buildIntelligentApps"
  // );w

  // if (!skipRemindInstallTeamsAgent) {
  //   void vscode.window
  //     .showInformationMessage(
  //       localize("teamstoolkit.handlers.askInstallTeamsAgent"),
  //       localize("teamstoolkit.handlers.askInstallTeamsAgent.install"),
  //       localize("teamstoolkit.handlers.askInstallTeamsAgent.confirmInstall")
  //     )
  //     .then(async (selection) => {
  //       await handleInstallTeamsAgentSelection(selection, triggerFromProperty);
  //     });
  // }
  // const data = { category: "TeamsDevApp.ms-teams-vscode-extension#buildIntelligentApps", step: "intelligentappresources" }
  // const tt =  await vscode.commands.executeCommand(
  //   "workbench.action.openWalkthrough",
  //   data
  // );
  await vscode.commands.executeCommand("workbench.action.openWalkthrough", {
    category: "TeamsDevApp.ms-teams-vscode-extension#teamsAgentGetStarted",
  });
  const isExtensionInstalled = githubCopilotInstalled();
  if (isExtensionInstalled) {
    VsCodeLogInstance.info(
      util.format(localize("teamstoolkit.handlers.installAgent.output"), teamsAgentLink)
    );
    showOutputChannelHandler();
    // await vscode.commands.executeCommand('workbench.action.openWalkthrough', {
    //   walkthrough: 'TeamsDevApp.ms-teams-vscode-extension#buildIntelligentApps',
    //   skipWelcome: true,
    //   context: context // Pass the context dynamically
    // });
    // return await vscode.commands.executeCommand(
    //   "workbench.action.openWalkthrough",
    //   "TeamsDevApp.ms-teams-vscode-extension#buildIntelligentApps"
    // );
    return ok(null);
    //return await openGithubCopilotChat(query);
  } else {
    VsCodeLogInstance.info(
      util.format(
        localize("teamstoolkit.handlers.installCopilotAndAgent.output"),
        InstallCopilotChatLink,
        teamsAgentLink
      )
    );
    showOutputChannelHandler();

    return ok(null);

    // const maxRetry = 5;
    // const installRes = await installGithubCopilotChatExtension(
    //   triggerFromProperty[TelemetryProperty.TriggerFrom]
    // );
    // if (installRes.isOk()) {
    //   let checkCount = 0;
    //   let verifyExtensionInstalled = false;
    //   while (checkCount < maxRetry) {
    //     verifyExtensionInstalled = githubCopilotInstalled();
    //     if (!verifyExtensionInstalled) {
    //       await sleep(3000);
    //       checkCount++;
    //     } else {
    //       break;
    //     }
    //   }

    //   if (verifyExtensionInstalled) {
    //     await sleep(2000); // wait for extension activation
    //     return await openGithubCopilotChat(query);
    //   } else {
    //     const error = new SystemError(
    //       eventName,
    //       errorNames.CannotVerifyGithubCopilotChat,
    //       util.format(
    //         localize("teamstoolkit.handlers.verifyCopilotExtensionError", InstallCopilotChatLink)
    //       ),
    //       util.format(
    //         localize("teamstoolkit.handlers.verifyCopilotExtensionError", InstallCopilotChatLink)
    //       )
    //     );
    //     VsCodeLogInstance.error(error.message);
    //     return err(error);
    //   }
    // } else {
    //   return installRes;
    // }
  }
}

/**
 * Invokes GitHub Copilot Chat for creating new app or development questions.
 * @param args args
 * @returns Result
 */
export async function invokeTeamsAgent(args?: any[]): Promise<Result<null, FxError>> {
  const eventName = TelemetryEvent.InvokeTeamsAgent;
  const triggerFromProperty = getTriggerFromProperty(args);
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.InvokeTeamsAgentStart, triggerFromProperty);

  let query = "";
  switch (triggerFromProperty[TelemetryProperty.TriggerFrom]) {
    case TelemetryTriggerFrom.TreeView:
    case TelemetryTriggerFrom.CommandPalette:
      query =
        "@teamsapp Use this GitHub Copilot extension to ask questions about Teams app and agent development.";
      break;
    case TelemetryTriggerFrom.WalkThroughIntroduction:
      query = "@teamsapp What is notification bot in Teams?";
      break;
    case TelemetryTriggerFrom.WalkThroughCreate:
      query = "@teamsapp How to create notification bot with Teams Toolkit?";
      break;
    case TelemetryTriggerFrom.WalkThroughWhatIsNext:
      query =
        "@teamsapp How do I customize and extend the notification bot app template created by Teams Toolkit?";
      break;
    case TelemetryTriggerFrom.WalkThroughIntelligentAppsIntroduction:
      query = "@teamsapp What is declarative agent for Microsoft 365 Copilot?";
      break;
    case TelemetryTriggerFrom.WalkThroughIntelligentAppsCreate:
      query = "@teamsapp How to create declarative agent with Teams Toolkit?";
      break;
    default:
      query =
        "@teamsapp Write your own query message to find relevant templates or samples to build your Teams app and agent as per your description. E.g. @teamsapp create an AI assistant bot that can complete common tasks.";
  }

  const res = await invoke(query, eventName, triggerFromProperty);

  if (res.isErr()) {
    ExtTelemetry.sendTelemetryErrorEvent(eventName, res.error, triggerFromProperty);
  } else {
    ExtTelemetry.sendTelemetryEvent(eventName, {
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
      ...triggerFromProperty,
    });
  }
  return res;
}

/**
 * Invokes teams agent for troubleshooting based on selected text.
 * @param args
 * @returns Result
 */
export async function troubleshootSelectedText(args?: any[]): Promise<Result<null, FxError>> {
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
  const res = await invoke(query, eventName, triggerFromProperty);

  if (res.isErr()) {
    ExtTelemetry.sendTelemetryErrorEvent(eventName, res.error, triggerFromProperty);
  } else {
    ExtTelemetry.sendTelemetryEvent(eventName, {
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
      ...triggerFromProperty,
    });
  }
  return res;
}

/**
 * Invokes teams agent for troubleshooting current error.
 * @param args
 * @returns Result
 */
export async function troubleshootError(args?: any[]): Promise<Result<null, FxError>> {
  const eventName = TelemetryEvent.TroubleshootErrorFromNotification;
  if (!args || args.length !== 2) {
    // should never happen
    return ok(null);
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
  const res = await invoke(query, eventName, triggerFromProperty);

  if (res.isErr()) {
    ExtTelemetry.sendTelemetryErrorEvent(eventName, res.error, telemtryProperties);
  } else {
    ExtTelemetry.sendTelemetryEvent(eventName, {
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
      ...telemtryProperties,
    });
  }
  return res;
}

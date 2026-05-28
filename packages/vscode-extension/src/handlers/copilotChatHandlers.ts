// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as util from "util";
import * as vscode from "vscode";

import { FxError, Result, SystemError, err, ok } from "@microsoft/teamsfx-api";
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
      "openCopilotError",
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
  triggerFrom: TelemetryTriggerFrom
): Promise<Result<null, FxError>> {
  const startEventName = "install-copilot-chat-start";
  const eventName = "install-copilot-chat";
  const telemetryProperties = {
    [TelemetryProperty.TriggerFrom]: triggerFrom,
  };
  ExtTelemetry.sendTelemetryEvent(startEventName, telemetryProperties);
  try {
    const confirmRes = await vscode.window.showInformationMessage(
      localize("teamstoolkit.handlers.askInstallCopilot"),
      localize("teamstoolkit.handlers.askInstallCopilot.install")
    );

    if (confirmRes !== localize("teamstoolkit.handlers.askInstallCopilot.install")) {
      const error = new UserCancelError(eventName, "cancel");
      ExtTelemetry.sendTelemetryErrorEvent(eventName, error, telemetryProperties);
      return err(error);
    } else {
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
    }
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

export async function invokeTeamsAgent(args?: any[]): Promise<Result<null, FxError>> {
  const eventName = TelemetryEvent.InvokeTeamsAgent;
  const triggerFromProperty = getTriggerFromProperty(args);
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.InvokeTeamsAgentStart, triggerFromProperty);

  const query =
    triggerFromProperty["trigger-from"] === TelemetryTriggerFrom.TreeView ||
    triggerFromProperty["trigger-from"] === TelemetryTriggerFrom.CommandPalette
      ? "@teamsapp Use this GitHub Copilot extension to ask questions about Teams app and agent development."
      : "@teamsapp Write your own query message to find relevant templates or samples to build your Teams app and agent as per your description. E.g. @teamsapp create an AI assistant bot that can complete common tasks.";
  let res: Result<null, FxError>;

  const skipRemindInstallTeamsAgent = await globalStateGet(
    GlobalKey.DoNotRemindInstallTeamsAgent,
    false
  );
  if (!skipRemindInstallTeamsAgent) {
    void vscode.window
      .showInformationMessage(
        localize("teamstoolkit.handlers.askInstallTeamsAgent"),
        localize("teamstoolkit.handlers.askInstallTeamsAgent.install"),
        localize("teamstoolkit.handlers.askInstallTeamsAgent.confirmInstall")
      )
      .then(async (selection) => {
        await handleInstallTeamsAgentSelection(selection, triggerFromProperty);
      });
  }

  const isExtensionInstalled = githubCopilotInstalled();
  if (isExtensionInstalled) {
    VsCodeLogInstance.info(
      util.format(localize("teamstoolkit.handlers.installAgent.output"), teamsAgentLink)
    );
    showOutputChannelHandler();
    res = await openGithubCopilotChat(query);
  } else {
    VsCodeLogInstance.info(
      util.format(
        localize("teamstoolkit.handlers.installCopilotAndAgent.output"),
        InstallCopilotChatLink,
        teamsAgentLink
      )
    );
    showOutputChannelHandler();

    const maxRetry = 5;
    const installRes = await installGithubCopilotChatExtension(
      triggerFromProperty[TelemetryProperty.TriggerFrom]
    );
    if (installRes.isOk()) {
      let checkCount = 0;
      let verifyExtensionInstalled = false;
      while (checkCount < maxRetry) {
        verifyExtensionInstalled = githubCopilotInstalled();
        if (!verifyExtensionInstalled) {
          await sleep(3000);
          checkCount++;
        } else {
          break;
        }
      }

      if (verifyExtensionInstalled) {
        await sleep(2000); // wait for extension activation
        res = await openGithubCopilotChat(query);
      } else {
        const error = new SystemError(
          eventName,
          "CannotVerifyGithubCopilotChat",
          util.format(
            localize("teamstoolkit.handlers.verifyCopilotExtensionError", InstallCopilotChatLink)
          ),
          util.format(
            localize("teamstoolkit.handlers.verifyCopilotExtensionError", InstallCopilotChatLink)
          )
        );
        VsCodeLogInstance.error(error.message);
        res = err(error);
      }
    } else {
      res = installRes;
    }
  }
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

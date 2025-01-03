// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserError, SystemError, FxError, Result, err, ok } from "@microsoft/teamsfx-api";
import {
  isUserCancelError,
  ConcurrentError,
  featureFlagManager,
  FeatureFlags as CoreFeatureFlags,
} from "@microsoft/teamsfx-core";
import { Uri, commands, window } from "vscode";
import {
  RecommendedOperations,
  openTestToolMessage,
  openTestToolDisplayMessage,
} from "../debug/common/debugConstants";
import {
  setOutputTroubleshootNotificationCount,
  outputTroubleshootNotificationCount,
  workspaceUri,
} from "../globalVariables";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import { anonymizeFilePaths } from "../utils/fileSystemUtils";
import { localize } from "../utils/localizeUtils";
import { isTestToolEnabledProject } from "../utils/projectChecker";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetryTriggerFrom,
} from "../telemetry/extTelemetryEvents";
import VsCodeLogInstance from "../commonlib/log";
import { ExtensionSource, ExtensionErrors } from "./error";
import { MaximumNotificationOutputTroubleshootCount } from "../constants";
import { sleep } from "@microsoft/vscode-ui";
import * as util from "util";

export async function showError(e: UserError | SystemError) {
  let notificationMessage = e.displayMessage ?? e.message;
  const errorCode = `${e.source}.${e.name}`;
  const runTestTool = {
    title: localize("teamstoolkit.handlers.debugInTestTool"),
    run: async () => {
      ExtTelemetry.sendTelemetryEvent(TelemetryEvent.MessageDebugInTestTool);
      await commands.executeCommand("workbench.action.quickOpen", "debug Debug in Test Tool");
      return ok<unknown, FxError>(null);
    },
  };
  const recommendTestTool =
    e.recommendedOperation === RecommendedOperations.DebugInTestTool &&
    workspaceUri?.fsPath &&
    isTestToolEnabledProject(workspaceUri.fsPath);

  const shouldRecommendTeamsAgent = featureFlagManager.getBooleanValue(
    CoreFeatureFlags.ChatParticipantUIEntries
  );
  const troubleshootErrorWithTeamsAgentButton = {
    title: localize("teamstoolkit.commmands.teamsAgentResolve.title"),
    run: async () => {
      await commands.executeCommand(
        "fx-extension.teamsAgentTroubleshootError",
        TelemetryTriggerFrom.Notification,
        e
      );
    },
  };

  if (recommendTestTool) {
    const recommendTestToolMessage = openTestToolMessage();
    const recommendTestToolDisplayMessage = openTestToolDisplayMessage();
    e.message += ` ${recommendTestToolMessage}`;
    notificationMessage += ` ${recommendTestToolDisplayMessage}`;
  }
  if (isUserCancelError(e)) {
    return;
  } else if ("helpLink" in e && e.helpLink && typeof e.helpLink != "undefined") {
    const helpLinkUrl = Uri.parse(`${e.helpLink}`);
    const help = {
      title: localize("teamstoolkit.handlers.getHelp"),
      run: () => {
        ExtTelemetry.sendTelemetryEvent(TelemetryEvent.ClickGetHelp, {
          [TelemetryProperty.ErrorCode]: errorCode,
          [TelemetryProperty.ErrorMessage]: notificationMessage,
          [TelemetryProperty.HelpLink]: e.helpLink!,
        });
        void commands.executeCommand("vscode.open", helpLinkUrl);
      },
    };
    VsCodeLogInstance.error(`code:${errorCode}, message: ${e.message}\n Help link: ${e.helpLink}`);
    // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
    VsCodeLogInstance.debug(`Call stack: ${e.stack || e.innerError?.stack || ""}`);
    const buttons = recommendTestTool ? [runTestTool, help] : [help];
    if (shouldRecommendTeamsAgent) {
      buttons.push(troubleshootErrorWithTeamsAgentButton);
    }
    void window
      .showErrorMessage(`[${errorCode}]: ${notificationMessage}`, ...buttons)
      .then((button) => {
        if (button) button.run();
      });
  } else if (e instanceof SystemError) {
    const sysError = e;
    const path = "https://github.com/OfficeDev/TeamsFx/issues/new?";
    const param = `title=bug+report: ${errorCode}&body=${anonymizeFilePaths(
      e.message
    )}\n\nstack:\n${anonymizeFilePaths(e.stack)}\n\n${
      sysError.userData ? anonymizeFilePaths(sysError.userData) : ""
    }`;
    const issueLink = Uri.parse(`${path}${param}`);
    const issue = {
      title: localize("teamstoolkit.handlers.reportIssue"),
      run: () => {
        void commands.executeCommand("vscode.open", issueLink);
      },
    };

    const similarIssues = {
      title: localize("teamstoolkit.handlers.similarIssues"),
      run: async (): Promise<void> => {
        await commands.executeCommand("fx-extension.findSimilarIssue", errorCode);
      },
    };
    const buttons = recommendTestTool ? [runTestTool, issue] : [issue];
    if (shouldRecommendTeamsAgent) {
      if (buttons.length >= 2) {
        buttons.push(troubleshootErrorWithTeamsAgentButton);
        notificationMessage += util.format(
          localize("teamstoolkit.handlers.similarIssues.message"),
          errorCode
        );
        e.message += util.format(
          localize("teamstoolkit.handlers.similarIssues.message"),
          errorCode
        );
      } else {
        buttons.push(similarIssues);
        buttons.push(troubleshootErrorWithTeamsAgentButton);
      }
    } else {
      buttons.push(similarIssues);
    }
    VsCodeLogInstance.error(`code:${errorCode}, message: ${e.message}`);
    // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
    VsCodeLogInstance.debug(`Call stack: ${e.stack || e.innerError?.stack || ""}`);

    void window
      .showErrorMessage(`[${errorCode}]: ${notificationMessage}`, ...buttons)
      .then((button) => {
        if (button) button.run();
      });
  } else {
    if (!(e instanceof ConcurrentError)) {
      // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
      VsCodeLogInstance.debug(`Call stack: ${e.stack || e.innerError?.stack || ""}`);
      const buttons: {
        title: string;
        run: () => void;
      }[] = recommendTestTool ? [runTestTool] : [];
      if (shouldRecommendTeamsAgent) {
        buttons.push(troubleshootErrorWithTeamsAgentButton);
      }
      void window
        .showErrorMessage(`[${errorCode}]: ${notificationMessage}`, ...buttons)
        .then((button) => {
          if (button) button.run();
        });
    }
  }

  if (shouldRecommendTeamsAgent && !isUserCancelError(e)) {
    await notifyOutputTroubleshoot(errorCode);
  }
}

export function wrapError(e: Error): Result<null, FxError> {
  if (
    e instanceof UserError ||
    e instanceof SystemError ||
    (e.constructor &&
      e.constructor.name &&
      (e.constructor.name === "SystemError" || e.constructor.name === "UserError"))
  ) {
    return err(e as FxError);
  }
  return err(
    new SystemError({ error: e, source: ExtensionSource, name: ExtensionErrors.UnknwonError })
  );
}

export function isLoginFailureError(error: FxError): boolean {
  return !!error.message && error.message.includes("Cannot get user login information");
}

export async function notifyOutputTroubleshoot(errorCode: string) {
  if (outputTroubleshootNotificationCount < MaximumNotificationOutputTroubleshootCount) {
    await sleep(3000);
    ExtTelemetry.sendTelemetryEvent(TelemetryEvent.NotifyOutputTroubleshoot, {
      [TelemetryProperty.ErrorCode]: errorCode,
    });
    setOutputTroubleshootNotificationCount(outputTroubleshootNotificationCount + 1);
    const buttonMessage = localize("teamstoolkit.commmands.teamsAgentResolve.openOutputPanel");
    void window
      .showInformationMessage(
        localize("teamstoolkit.commmands.teamsAgentResolve.dialogContent"),
        buttonMessage
      )
      .then((selection) => {
        if (selection === buttonMessage) {
          ExtTelemetry.sendTelemetryEvent(TelemetryEvent.ClickToOpenOutputForTroubleshoot);
          void commands.executeCommand("fx-extension.showOutputChannel");
        }
      });
  }
}

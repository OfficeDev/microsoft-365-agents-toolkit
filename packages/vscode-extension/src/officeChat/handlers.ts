// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Correlator } from "@microsoft/teamsfx-core";
import { ConstantString } from "@microsoft/teamsfx-core/build/src/common/constants";
import fs from "fs-extra";
import os from "os";
import path from "path";
import {
  CancellationToken,
  ChatContext,
  ChatFollowup,
  ChatRequest,
  ChatResponseStream,
  ChatResultFeedback,
  LanguageModelChatMessage,
  LanguageModelChatMessageRole,
  ProviderResult,
  QuickPickItem,
  Uri,
  window,
} from "vscode";
import followupProvider from "../chat/followupProvider";
import { ITelemetryData } from "../chat/types";
import { verbatimCopilotInteraction } from "../chat/utils";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetryTriggerFrom,
} from "../telemetry/extTelemetryEvents";
import { localize } from "../utils/localizeUtils";
import { openOfficeDevFolder } from "../utils/workspaceUtils";
import officeCreateCommandHandler from "./commands/create/officeCreateCommandHandler";
import generatecodeCommandHandler from "./commands/generatecode/generatecodeCommandHandler";
import officeNextStepCommandHandler from "./commands/nextStep/officeNextstepCommandHandler";
import { OfficeChatCommand, officeChatParticipantId } from "./consts";
import { defaultOfficeSystemPrompt } from "./officePrompts";
import { OfficeChatTelemetryData } from "./telemetry";
import { ICopilotChatOfficeResult } from "./types";

export function officeChatRequestHandler(
  request: ChatRequest,
  context: ChatContext,
  response: ChatResponseStream,
  token: CancellationToken
): ProviderResult<ICopilotChatOfficeResult> {
  followupProvider.clearFollowups();
  if (request.command == OfficeChatCommand.Create) {
    return officeCreateCommandHandler(request, context, response, token);
  } else if (request.command == OfficeChatCommand.GenerateCode) {
    return generatecodeCommandHandler(request, context, response, token);
  } else if (request.command == OfficeChatCommand.NextStep) {
    return officeNextStepCommandHandler(request, context, response, token);
  } else {
    return officeDefaultHandler(request, context, response, token);
  }
}

async function officeDefaultHandler(
  request: ChatRequest,
  context: ChatContext,
  response: ChatResponseStream,
  token: CancellationToken
): Promise<ICopilotChatOfficeResult> {
  const officeChatTelemetryData = OfficeChatTelemetryData.createByParticipant(
    officeChatParticipantId,
    ""
  );
  ExtTelemetry.sendTelemetryEvent(
    TelemetryEvent.CopilotChatStart,
    officeChatTelemetryData.properties
  );

  if (!request.prompt) {
    throw new Error(`
Please specify a question when using this command.

Usage: @office Ask questions about Office Add-ins development.`);
  }
  const messages = [
    defaultOfficeSystemPrompt(),
    new LanguageModelChatMessage(LanguageModelChatMessageRole.User, request.prompt),
  ];
  officeChatTelemetryData.chatMessages.push(...messages);
  await verbatimCopilotInteraction("copilot-gpt-4", messages, response, token);
  const followUps: ChatFollowup[] = [
    {
      label: "@office /create an Excel hello world add-in",
      command: "create",
      prompt: "an Excel hello world add-in",
    },
  ];
  followupProvider.addFollowups(followUps);

  officeChatTelemetryData.markComplete();
  ExtTelemetry.sendTelemetryEvent(
    TelemetryEvent.CopilotChat,
    officeChatTelemetryData.properties,
    officeChatTelemetryData.measurements
  );
  return { metadata: { command: undefined, requestId: officeChatTelemetryData.requestId } };
}

export async function chatCreateOfficeProjectCommandHandler(
  folder: string,
  requestId: string,
  matchResultInfo: string,
  appId: string
) {
  const officeChatTelemetryData = OfficeChatTelemetryData.get(requestId);
  if (officeChatTelemetryData) {
    ExtTelemetry.sendTelemetryEvent(
      TelemetryEvent.CopilotChatClickButton,
      {
        ...officeChatTelemetryData.properties,
        [TelemetryProperty.CopilotMatchResultType]: matchResultInfo,
      },
      officeChatTelemetryData.measurements
    );
  }
  // Let user choose the project folder
  let dstPath = "";
  let folderChoice: QuickPickItem | undefined = undefined;
  const defaultFolder = path.join(os.homedir(), ConstantString.RootFolder);
  folderChoice = await window.showQuickPick([
    {
      label: localize("teamstoolkit.qm.defaultFolder"),
      description: defaultFolder,
    },
    {
      label: localize("teamstoolkit.qm.browse"),
    },
  ]);
  if (!folderChoice) {
    return;
  }
  if (folderChoice.label === localize("teamstoolkit.qm.defaultFolder")) {
    dstPath = defaultFolder;
  } else {
    const customFolder = await window.showOpenDialog({
      title: localize("teamstoolkit.chatParticipants.officeAddIn.create.selectFolder.title"),
      openLabel: localize("teamstoolkit.chatParticipants.officeAddIn.create.selectFolder.label"),
      canSelectFiles: false,
      canSelectFolders: true,
      canSelectMany: false,
    });
    if (!customFolder) {
      return;
    }
    dstPath = customFolder[0].fsPath;
  }
  try {
    let workDir = path.join(dstPath, appId);
    let suffix = 1;
    while (fs.pathExistsSync(workDir) && fs.readdirSync(workDir).length > 0) {
      workDir = path.join(dstPath, `${appId}_${suffix++}`);
    }
    fs.ensureDirSync(workDir);
    await fs.copy(folder, workDir);
    await openOfficeDevFolder(Uri.file(workDir), true);
  } catch (error) {
    console.error("Error copying files:", error);
    void window.showErrorMessage(
      localize("teamstoolkit.chatParticipants.officeAddIn.create.failToCreate")
    );
  }
}

export function handleOfficeFeedback(e: ChatResultFeedback): void {
  const result = e.result as ICopilotChatOfficeResult;
  const telemetryData: ITelemetryData = {
    properties: {
      [TelemetryProperty.CopilotChatRequestId]: result.metadata?.requestId ?? "",
      [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.CopilotChat,
      [TelemetryProperty.CopilotChatCommand]: result.metadata?.command ?? "",
      [TelemetryProperty.CorrelationId]: Correlator.getId(),
      [TelemetryProperty.HostType]:
        OfficeChatTelemetryData.get(result.metadata?.requestId ?? "")?.properties[
          TelemetryProperty.HostType
        ] ?? "",
      [TelemetryProperty.CopilotChatRelatedSampleName]:
        OfficeChatTelemetryData.get(result.metadata?.requestId ?? "")?.properties[
          TelemetryProperty.CopilotChatRelatedSampleName
        ] ?? "",
    },
    measurements: {
      [TelemetryProperty.CopilotChatFeedbackHelpful]: e.kind,
    },
  };

  ExtTelemetry.sendTelemetryEvent(
    TelemetryEvent.CopilotChatFeedback,
    telemetryData.properties,
    telemetryData.measurements
  );
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { isValidOfficeAddInProject } from "@microsoft/teamsfx-core";
import {
  CancellationToken,
  ChatContext,
  ChatFollowup,
  ChatRequest,
  ChatResponseStream,
  LanguageModelChatMessage,
  LanguageModelChatMessageRole,
} from "vscode";
import { workspaceUri } from "../../../globalVariables";
import { ExtTelemetry } from "../../../telemetry/extTelemetry";
import { TelemetryEvent } from "../../../telemetry/extTelemetryEvents";
import { CHAT_EXECUTE_COMMAND_ID } from "../../../chat/consts";
import { OfficeChatCommand, officeChatParticipantId } from "../../consts";
import followupProvider from "../../../chat/followupProvider";
import { officeSteps } from "./officeSteps";
import { OfficeWholeStatus } from "./types";
import { getWholeStatus } from "./status";
import { localize } from "../../../utils/localizeUtils";
import { ICopilotChatOfficeResult } from "../../types";
import { NextStep } from "../../../chat/commands/nextstep/types";
import { describeOfficeStepSystemPrompt } from "../../officePrompts";
import { getCopilotResponseAsString } from "../../../chat/utils";
import { IChatTelemetryData } from "../../../chat/types";
import { OfficeChatTelemetryBlockReasonEnum, OfficeChatTelemetryData } from "../../telemetry";

export default async function officeNextStepCommandHandler(
  request: ChatRequest,
  context: ChatContext,
  response: ChatResponseStream,
  token: CancellationToken
): Promise<ICopilotChatOfficeResult> {
  const officeChatTelemetryData = OfficeChatTelemetryData.createByParticipant(
    officeChatParticipantId,
    OfficeChatCommand.NextStep
  );
  ExtTelemetry.sendTelemetryEvent(
    TelemetryEvent.CopilotChatStart,
    officeChatTelemetryData.properties
  );

  if (request.prompt) {
    officeChatTelemetryData.setTimeToFirstToken();
    response.markdown(localize("teamstoolkit.chatParticipants.officeAddIn.nextStep.promptAnswer"));
    officeChatTelemetryData.setBlockReason(OfficeChatTelemetryBlockReasonEnum.UnsupportedInput);
    officeChatTelemetryData.markComplete("fail");
    ExtTelemetry.sendTelemetryEvent(
      TelemetryEvent.CopilotChat,
      officeChatTelemetryData.properties,
      officeChatTelemetryData.measurements
    );
    return {
      metadata: {
        command: OfficeChatCommand.NextStep,
        requestId: officeChatTelemetryData.requestId,
      },
    };
  }

  const workspace = workspaceUri?.fsPath;
  const officeAddInApp = isValidOfficeAddInProject(workspace) ? workspace : undefined;
  const status: OfficeWholeStatus = await getWholeStatus(officeAddInApp);
  const steps = officeSteps()
    .filter((s) => s.condition(status))
    .sort((a, b) => a.priority - b.priority);
  if (steps.length > 1) {
    officeChatTelemetryData.setTimeToFirstToken();
    response.markdown("Here are the next steps you can do:\n");
  }
  for (let index = 0; index < Math.min(3, steps.length); index++) {
    const s = steps[index];
    if (s.description instanceof Function) {
      s.description = s.description(status);
    }
    const stepDescription = s.description;
    const title = s.docLink ? `[${s.title}](${s.docLink})` : s.title;
    if (steps.length > 1) {
      response.markdown(`${index + 1}. ${title}: ${stepDescription}\n`);
    } else {
      officeChatTelemetryData.setTimeToFirstToken();
      response.markdown(`${title}: ${stepDescription}\n`);
    }
    s.commands.forEach((c) => {
      if (c.command === CHAT_EXECUTE_COMMAND_ID) {
        c.arguments!.splice(1, 0, officeChatTelemetryData.requestId);
      }
      response.button(c);
    });
  }
  const followUps: ChatFollowup[] = [];
  steps.forEach((s) => {
    followUps.push(...s.followUps);
  });
  followupProvider.addFollowups(followUps);

  officeChatTelemetryData.markComplete();
  ExtTelemetry.sendTelemetryEvent(
    TelemetryEvent.CopilotChat,
    officeChatTelemetryData.properties,
    officeChatTelemetryData.measurements
  );

  return {
    metadata: {
      command: OfficeChatCommand.NextStep,
      requestId: officeChatTelemetryData.requestId,
    },
  };
}

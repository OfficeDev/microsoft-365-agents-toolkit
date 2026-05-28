// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { isValidProject } from "@microsoft/teamsfx-core";
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
import { CHAT_EXECUTE_COMMAND_ID, TeamsChatCommand, chatParticipantId } from "../../consts";
import followupProvider from "../../followupProvider";
import { describeStepSystemPrompt } from "../../prompts";
import { ChatTelemetryData } from "../../telemetry";
import { IChatTelemetryData, ICopilotChatResult } from "../../types";
import { getCopilotResponseAsString } from "../../utils";
import { getWholeStatus } from "./status";
import { allSteps } from "./steps";
import { NextStep, WholeStatus } from "./types";
import { localize } from "../../../utils/localizeUtils";

export default async function nextStepCommandHandler(
  request: ChatRequest,
  context: ChatContext,
  response: ChatResponseStream,
  token: CancellationToken
): Promise<ICopilotChatResult> {
  const chatTelemetryData = ChatTelemetryData.createByParticipant(
    chatParticipantId,
    TeamsChatCommand.NextStep
  );
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.CopilotChatStart, chatTelemetryData.properties);

  if (request.prompt) {
    response.markdown(localize("teamstoolkit.chatParticipants.nextStep.noPromptAnswer"));
    chatTelemetryData.markComplete("unsupportedPrompt");
    ExtTelemetry.sendTelemetryEvent(
      TelemetryEvent.CopilotChat,
      chatTelemetryData.properties,
      chatTelemetryData.measurements
    );
    return {
      metadata: {
        command: TeamsChatCommand.NextStep,
        requestId: chatTelemetryData.requestId,
      },
    };
  }

  const workspace = workspaceUri?.fsPath;
  const teamsApp = isValidProject(workspace) ? workspace : undefined;
  const status: WholeStatus = await getWholeStatus(teamsApp);
  const steps = allSteps()
    .filter((s) => s.condition(status))
    .sort((a, b) => a.priority - b.priority);
  if (steps.length > 1) {
    response.markdown("Here are the next steps you can do:\n");
  }
  for (let index = 0; index < Math.min(3, steps.length); index++) {
    const s = steps[index];
    if (s.description instanceof Function) {
      s.description = s.description(status);
    }
    const stepDescription = await describeStep(s, token, chatTelemetryData);
    const title = s.docLink ? `[${s.title}](${s.docLink})` : s.title;
    if (steps.length > 1) {
      response.markdown(`${index + 1}. ${title}: ${stepDescription}\n`);
    } else {
      response.markdown(`${title}: ${stepDescription}\n`);
    }
    s.commands.forEach((c) => {
      if (c.command === CHAT_EXECUTE_COMMAND_ID) {
        c.arguments!.splice(1, 0, chatTelemetryData.requestId);
      }
      response.button(c);
    });
  }
  const followUps: ChatFollowup[] = [];
  steps.forEach((s) => {
    const ids = followUps.map((f) => `${f.label ?? ""} ${f.command ?? ""} ${f.prompt}`);
    followUps.push(
      ...s.followUps.filter((f) => !ids.includes(`${f.label ?? ""} ${f.command ?? ""} ${f.prompt}`))
    );
  });
  followupProvider.addFollowups(followUps);

  chatTelemetryData.markComplete();
  ExtTelemetry.sendTelemetryEvent(
    TelemetryEvent.CopilotChat,
    chatTelemetryData.properties,
    chatTelemetryData.measurements
  );

  return {
    metadata: {
      command: TeamsChatCommand.NextStep,
      requestId: chatTelemetryData.requestId,
    },
  };
}

export async function describeStep(
  step: NextStep,
  token: CancellationToken,
  telemetryMetadata: IChatTelemetryData
): Promise<string> {
  const messages = [
    describeStepSystemPrompt(),
    new LanguageModelChatMessage(
      LanguageModelChatMessageRole.User,
      `The content is '${JSON.stringify({
        description: step.description as string,
      })}'.`
    ),
  ];

  telemetryMetadata.chatMessages.push(...messages);
  return await getCopilotResponseAsString("copilot-gpt-3.5-turbo", messages, token);
}

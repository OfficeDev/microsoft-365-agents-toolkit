// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  CancellationToken,
  ChatContext,
  ChatRequest,
  ChatResponseStream,
  ChatResultFeedback,
  LanguageModelChatMessage,
  LanguageModelChatMessageRole,
  ProviderResult,
  Uri,
  commands,
  env,
} from "vscode";

import * as uuid from "uuid";

import { FxError, Result } from "@microsoft/teamsfx-api";
import { Correlator } from "@microsoft/teamsfx-core";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetryTriggerFrom,
} from "../telemetry/extTelemetryEvents";
import createCommandHandler from "./commands/create/createCommandHandler";
import nextStepCommandHandler from "./commands/nextstep/nextstepCommandHandler";
import { TeamsChatCommand, chatParticipantId } from "./consts";
import followupProvider from "./followupProvider";
import { defaultSystemPrompt } from "./prompts";
import { ChatTelemetryData } from "./telemetry";
import { ICopilotChatResult, ITelemetryData } from "./types";
import { verbatimCopilotInteraction } from "./utils";
import { CommandKey } from "../constants";

export function chatRequestHandler(
  request: ChatRequest,
  context: ChatContext,
  response: ChatResponseStream,
  token: CancellationToken
): ProviderResult<ICopilotChatResult> {
  // Matching chat commands in the package.json
  followupProvider.clearFollowups();
  if (request.command == TeamsChatCommand.Create) {
    return createCommandHandler(request, context, response, token);
  } else if (request.command == TeamsChatCommand.NextStep) {
    return nextStepCommandHandler(request, context, response, token);
  } else {
    return defaultHandler(request, context, response, token);
  }
}

async function defaultHandler(
  request: ChatRequest,
  context: ChatContext,
  response: ChatResponseStream,
  token: CancellationToken
): Promise<ICopilotChatResult> {
  const chatTelemetryData = ChatTelemetryData.createByParticipant(chatParticipantId, "");
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.CopilotChatStart, chatTelemetryData.properties);

  if (!request.prompt) {
    throw new Error(`
Please specify a question when using this command.

Usage: @teams Ask questions about Teams Development"`);
  }
  const messages = [
    defaultSystemPrompt(),
    new LanguageModelChatMessage(LanguageModelChatMessageRole.User, request.prompt),
  ];
  chatTelemetryData.chatMessages.push(...messages);
  await verbatimCopilotInteraction("copilot-gpt-4", messages, response, token);

  chatTelemetryData.markComplete();
  ExtTelemetry.sendTelemetryEvent(
    TelemetryEvent.CopilotChat,
    chatTelemetryData.properties,
    chatTelemetryData.measurements
  );
  return { metadata: { command: undefined, requestId: chatTelemetryData.requestId } };
}

export async function chatExecuteCommandHandler(
  command: string,
  requestId: string,
  ...args: unknown[]
): Promise<Result<unknown, FxError>> {
  const chatTelemetryData = ChatTelemetryData.get(requestId);
  const correlationId = uuid.v4();
  if (chatTelemetryData) {
    ExtTelemetry.sendTelemetryEvent(
      TelemetryEvent.CopilotChatClickButton,
      {
        ...chatTelemetryData.properties,
        [TelemetryProperty.CopilotChatRunCommandId]: command,
        [TelemetryProperty.CorrelationId]: correlationId,
      },
      chatTelemetryData.measurements
    );
  }
  if (Object.values(CommandKey).includes(command as CommandKey)) {
    return await commands.executeCommand<Result<unknown, FxError>>(
      command,
      correlationId,
      TelemetryTriggerFrom.CopilotChat,
      ...args
    );
  }
  return await commands.executeCommand(command, ...args);
}

export async function openUrlCommandHandler(url: string) {
  await env.openExternal(Uri.parse(url));
}

export function handleFeedback(e: ChatResultFeedback): void {
  const result = e.result as ICopilotChatResult;
  const telemetryData: ITelemetryData = {
    properties: {
      [TelemetryProperty.CopilotChatRequestId]: result.metadata?.requestId ?? "",
      [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.CopilotChat,
      [TelemetryProperty.CopilotChatCommand]: result.metadata?.command ?? "",
      [TelemetryProperty.CorrelationId]: Correlator.getId(),
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

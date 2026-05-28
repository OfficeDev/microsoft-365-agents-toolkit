// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { LanguageModelChatMessage } from "vscode";
import { countMessagesTokens } from "./utils";
import { IChatTelemetryData, ITelemetryData } from "./types";
import { Correlator, getUuid } from "@microsoft/teamsfx-core";
import {
  TelemetryProperty,
  TelemetrySuccess,
  TelemetryTriggerFrom,
} from "../telemetry/extTelemetryEvents";

export class ChatTelemetryData implements IChatTelemetryData {
  public static requestData: { [key: string]: ChatTelemetryData } = {};

  telemetryData: ITelemetryData;
  chatMessages: LanguageModelChatMessage[] = [];
  command: string;
  requestId: string;
  startTime: number;
  // participant name
  participantId: string;
  // The location at which the chat is happening.
  hasComplete = false;

  get properties(): { [key: string]: string } {
    return this.telemetryData.properties;
  }

  get measurements(): { [key: string]: number } {
    return this.telemetryData.measurements;
  }

  constructor(command: string, requestId: string, startTime: number, participantId: string) {
    this.command = command;
    this.requestId = requestId;
    this.startTime = startTime;
    this.participantId = participantId;

    const telemetryData: ITelemetryData = { properties: {}, measurements: {} };
    telemetryData.properties[TelemetryProperty.CopilotChatCommand] = command;
    telemetryData.properties[TelemetryProperty.CopilotChatRequestId] = this.requestId;
    // currently only triggerd by copilot chat
    telemetryData.properties[TelemetryProperty.TriggerFrom] = TelemetryTriggerFrom.CopilotChat;
    telemetryData.properties[TelemetryProperty.CorrelationId] = Correlator.getId();
    telemetryData.properties[TelemetryProperty.CopilotChatParticipantId] = participantId;
    // The value of properties must be string type.
    this.telemetryData = telemetryData;

    ChatTelemetryData.requestData[requestId] = this;
  }

  static createByParticipant(participantId: string, command: string) {
    const requestId = getUuid();
    const startTime = Date.now();
    return new ChatTelemetryData(command, requestId, startTime, participantId);
  }

  static get(requestId: string): ChatTelemetryData | undefined {
    return ChatTelemetryData.requestData[requestId];
  }

  chatMessagesTokenCount(): number {
    return countMessagesTokens(this.chatMessages);
  }

  extendBy(properties?: { [key: string]: string }, measurements?: { [key: string]: number }) {
    this.telemetryData.properties = { ...this.telemetryData.properties, ...properties };
    this.telemetryData.measurements = { ...this.telemetryData.measurements, ...measurements };
  }

  markComplete(completeType: "success" | "unsupportedPrompt" = "success") {
    if (!this.hasComplete) {
      this.telemetryData.properties[TelemetryProperty.Success] = TelemetrySuccess.Yes;
      this.telemetryData.properties[TelemetryProperty.CopilotChatCompleteType] = completeType;
      this.telemetryData.measurements[TelemetryProperty.CopilotChatTimeToComplete] =
        Date.now() - this.startTime;
      this.telemetryData.measurements[TelemetryProperty.CopilotChatTokenCount] =
        this.chatMessagesTokenCount();
      this.hasComplete = true;
    }
  }
}

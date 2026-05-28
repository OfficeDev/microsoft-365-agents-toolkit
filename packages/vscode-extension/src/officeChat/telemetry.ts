// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { LanguageModelChatMessage } from "vscode";
import { IChatTelemetryData, ITelemetryData } from "../chat/types";
import { Correlator, getUuid } from "@microsoft/teamsfx-core";
import { countMessagesTokens } from "../chat/utils";
import {
  TelemetryProperty,
  TelemetrySuccess,
  TelemetryTriggerFrom,
} from "../telemetry/extTelemetryEvents";

export enum OfficeChatTelemetryBlockReasonEnum {
  RAI = "RAI",
  OffTopic = "Off Topic",
  UnsupportedInput = "Unsupported Input",
  LanguageModelError = "LanguageModel Error",
  PlannerFailure = "Planner Failure",
}
export class OfficeChatTelemetryData implements IChatTelemetryData {
  public static requestData: { [key: string]: OfficeChatTelemetryData } = {};

  telemetryData: ITelemetryData;
  chatMessages: LanguageModelChatMessage[] = [];
  responseChatMessages: LanguageModelChatMessage[] = [];
  command: string;
  requestId: string;
  startTime: number;
  hostType: string;
  relatedSampleName: string;
  timeToFirstToken: number;
  blockReason?: string;
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
    this.hostType = "";
    this.relatedSampleName = "";
    this.timeToFirstToken = -1;

    const telemetryData: ITelemetryData = { properties: {}, measurements: {} };
    telemetryData.properties[TelemetryProperty.CopilotChatCommand] = command;
    telemetryData.properties[TelemetryProperty.CopilotChatRequestId] = this.requestId;
    // currently only triggerd by copilot chat
    telemetryData.properties[TelemetryProperty.TriggerFrom] = TelemetryTriggerFrom.CopilotChat;
    telemetryData.properties[TelemetryProperty.CorrelationId] = Correlator.getId();
    telemetryData.properties[TelemetryProperty.CopilotChatParticipantId] = participantId;
    // The value of properties must be string type.
    this.telemetryData = telemetryData;

    OfficeChatTelemetryData.requestData[requestId] = this;
  }

  static createByParticipant(participantId: string, command: string) {
    const requestId = getUuid();
    const startTime = performance.now();
    return new OfficeChatTelemetryData(command, requestId, startTime, participantId);
  }

  static get(requestId: string): OfficeChatTelemetryData | undefined {
    return OfficeChatTelemetryData.requestData[requestId];
  }

  setHostType(hostType: string) {
    this.hostType = hostType;
  }

  setRelatedSampleName(relatedSampleName: string) {
    this.relatedSampleName = relatedSampleName;
  }

  setTimeToFirstToken(t0?: DOMHighResTimeStamp) {
    if (t0) {
      this.timeToFirstToken = (t0 - this.startTime) / 1000;
    } else {
      this.timeToFirstToken = (performance.now() - this.startTime) / 1000;
    }
  }

  setBlockReason(blockReason: string) {
    this.blockReason = blockReason;
  }

  chatMessagesTokenCount(): number {
    return countMessagesTokens(this.chatMessages);
  }

  responseChatMessagesTokenCount(): number {
    return countMessagesTokens(this.responseChatMessages);
  }

  extendBy(properties?: { [key: string]: string }, measurements?: { [key: string]: number }) {
    this.telemetryData.properties = { ...this.telemetryData.properties, ...properties };
    this.telemetryData.measurements = { ...this.telemetryData.measurements, ...measurements };
  }

  markComplete(completeType: "success" | "fail" = "success") {
    if (!this.hasComplete) {
      this.telemetryData.properties[TelemetryProperty.Success] = TelemetrySuccess.Yes;
      this.telemetryData.properties[TelemetryProperty.CopilotChatCompleteType] = completeType;
      if (this.blockReason && this.blockReason !== "") {
        this.telemetryData.properties[TelemetryProperty.CopilotChatBlockReason] = this.blockReason;
      }
      this.telemetryData.properties[TelemetryProperty.HostType] = this.hostType;
      this.telemetryData.properties[TelemetryProperty.CopilotChatRelatedSampleName] =
        this.relatedSampleName;
      this.telemetryData.measurements[TelemetryProperty.CopilotChatTimeToFirstToken] =
        this.timeToFirstToken;
      this.telemetryData.measurements[TelemetryProperty.CopilotChatTimeToComplete] =
        (performance.now() - this.startTime) / 1000;
      this.telemetryData.measurements[TelemetryProperty.CopilotChatRequestToken] =
        this.chatMessagesTokenCount();
      this.telemetryData.measurements[TelemetryProperty.CopilotChatResponseToken] =
        this.responseChatMessagesTokenCount();
      this.telemetryData.measurements[TelemetryProperty.CopilotChatRequestTokenPerSecond] =
        this.telemetryData.measurements[TelemetryProperty.CopilotChatRequestToken] /
        this.telemetryData.measurements[TelemetryProperty.CopilotChatTimeToComplete];
      this.telemetryData.measurements[TelemetryProperty.CopilotChatResponseTokenPerSecond] =
        this.telemetryData.measurements[TelemetryProperty.CopilotChatResponseToken] /
        this.telemetryData.measurements[TelemetryProperty.CopilotChatTimeToComplete];
      this.hasComplete = true;
    }
  }
}

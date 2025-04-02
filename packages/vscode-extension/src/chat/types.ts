// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { LanguageModelChatMessage } from "vscode";

export interface ITelemetryData {
  properties: { [key: string]: string };
  measurements: { [key: string]: number };
}

export interface IChatTelemetryData {
  telemetryData: ITelemetryData;
  chatMessages: LanguageModelChatMessage[];
  command: string;
  requestId: string;
  startTime: number;

  chatMessagesTokenCount: () => number;
  get properties(): { [key: string]: string };
  get measurements(): { [key: string]: number };
}

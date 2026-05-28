// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { ChatResult } from "vscode";
import { OfficeChatCommand } from "./consts";

export interface ICopilotChatOfficeResultMetadata {
  readonly command: OfficeChatCommand | undefined;
  readonly requestId: string;
}

export interface ICopilotChatOfficeResult extends ChatResult {
  readonly metadata?: ICopilotChatOfficeResultMetadata;
}

export type OfficeProjectInfo = {
  path: string;
  host: string;
};

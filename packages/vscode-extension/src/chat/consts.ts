// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ChatFollowup } from "vscode";

export const chatParticipantId = "ms-teams-vscode-extension.teams";

export const CHAT_EXECUTE_COMMAND_ID = "fx-extension.chat.executeCommand";
export const CHAT_OPENURL_COMMAND_ID = "fx-extension.chat.openUrlCommand";

export const enum TeamsChatCommand {
  Create = "create",
  NextStep = "nextstep",
}

export const DefaultNextStep: ChatFollowup = {
  prompt: "",
  command: "nextstep",
  label: "What should I do next?",
};

// for counting message token, refer to vscode copilot solution
/**
 * BaseTokensPerCompletion is the minimum tokens for a completion request.
 * Replies are primed with <|im_start|>assistant<|message|>, so these tokens represent the
 * special token and the role name.
 */
export const BaseTokensPerCompletion = 3;
/*
 * Each GPT 3.5 / GPT 4 message comes with 3 tokens per message due to special characters
 */
export const BaseTokensPerMessage = 3;
/*
 * Since gpt-3.5-turbo-0613 each name costs 1 token
 */
export const BaseTokensPerName = 1;

export const IsChatParticipantEnabled = true;

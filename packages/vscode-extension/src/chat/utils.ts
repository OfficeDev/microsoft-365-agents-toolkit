// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CancellationToken, ChatResponseStream, LanguageModelChatMessage, lm } from "vscode";

import { sampleProvider } from "@microsoft/teamsfx-core";
import { BaseTokensPerCompletion, BaseTokensPerMessage, BaseTokensPerName } from "./consts";
import { Tokenizer } from "./tokenizer";

export async function verbatimCopilotInteraction(
  model: "copilot-gpt-3.5-turbo" | "copilot-gpt-4",
  messages: LanguageModelChatMessage[],
  response: ChatResponseStream,
  token: CancellationToken
) {
  const [vendor, family] = model.split(/-(.*)/s);
  const chatModels = await lm.selectChatModels({ vendor, family });
  const familyMatch = chatModels?.find((chatModel) => chatModel.family === family);
  if (!familyMatch) {
    throw new Error("No chat models available for the specified family");
  }
  const chatResponse = await familyMatch.sendRequest(messages, {}, token);
  for await (const fragment of chatResponse.text) {
    response.markdown(fragment);
  }
}

export async function getCopilotResponseAsString(
  model: "copilot-gpt-3.5-turbo" | "copilot-gpt-4",
  messages: LanguageModelChatMessage[],
  token: CancellationToken
): Promise<string> {
  const [vendor, family] = model.split(/-(.*)/s);
  const chatModels = await lm.selectChatModels({ vendor, family });
  const familyMatch = chatModels?.find((chatModel) => chatModel.family === family);
  if (!familyMatch) {
    throw new Error("No chat models available for the specified family");
  }
  const chatResponse = await familyMatch.sendRequest(messages, {}, token);
  let response = "";
  for await (const fragment of chatResponse.text) {
    response += fragment;
  }
  return response;
}

export async function getSampleDownloadUrlInfo(sampleId: string) {
  const sampleCollection = await sampleProvider.SampleCollection;
  const sample = sampleCollection.samples.find((sample) => sample.id === sampleId);
  if (!sample) {
    throw new Error("Sample not found");
  }
  return sample.downloadUrlInfo;
}

// count message token for GPT3.5 and GPT4 message
// refer to vscode copilot tokenizer solution
export function countMessageTokens(message: LanguageModelChatMessage): number {
  let numTokens = BaseTokensPerMessage;
  const tokenizer = Tokenizer.getInstance();
  for (const [key, value] of Object.entries(message)) {
    if (!value || key === "role") {
      continue;
    }
    numTokens += tokenizer.tokenLength(value);
    if (key === "name") {
      numTokens += BaseTokensPerName;
    }
  }
  return numTokens;
}

export function countMessagesTokens(messages: LanguageModelChatMessage[]): number {
  let numTokens = 0;
  for (const message of messages) {
    numTokens += countMessageTokens(message);
  }
  numTokens += BaseTokensPerCompletion;
  return numTokens;
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CancellationToken, ChatResponseStream, LanguageModelChatMessage } from "vscode";
import { Spec } from "./spec";
import { ExecutionResultEnum } from "./executionResultEnum";

export interface ISkill {
  name: string | undefined;
  capability: string | undefined;
  canInvoke: (spec: Spec) => boolean;
  invoke: (
    languageModel: LanguageModelChatMessage,
    response: ChatResponseStream,
    token: CancellationToken,
    spec: Spec
  ) => Promise<{ result: ExecutionResultEnum; spec: Spec }>;
}

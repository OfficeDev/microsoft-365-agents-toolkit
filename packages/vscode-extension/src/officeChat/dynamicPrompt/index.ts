// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { LanguageModelChatMessage, LanguageModelChatMessageRole } from "vscode";
import { commonTemplates } from "./formats/common";
import { buildDynamicPromptInternal } from "./utils/buildDynamicPrompt";
import { IDynamicPromptFormat, MessageRole } from "./utils/types";

export interface IDynamicPrompt {
  messages: LanguageModelChatMessage[];
  version: string;
}

export function buildDynamicPrompt<T>(format: IDynamicPromptFormat<T>, args: T): IDynamicPrompt {
  try {
    const messages = format.messages.map((messageFormat) => {
      const { role, entryTemplate } = messageFormat;

      const prompt = buildDynamicPromptInternal(`templates.${entryTemplate}`, {
        args,
        common: commonTemplates,
        presets: format.presets,
        templates: format.templates,
      });

      return createMessage(role, prompt);
    });

    return {
      messages,
      version: format.version,
    };
  } catch (e) {
    throw e;
  }
}

function createMessage(role: MessageRole, prompt: string): LanguageModelChatMessage {
  switch (role) {
    case "system":
    case "user":
      return new LanguageModelChatMessage(LanguageModelChatMessageRole.User, prompt);
    case "assistant":
      return new LanguageModelChatMessage(LanguageModelChatMessageRole.Assistant, prompt);
  }
}

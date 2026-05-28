// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  CancellationToken,
  ChatResponseStream,
  LanguageModelChatMessage,
  LanguageModelChatMessageRole,
} from "vscode";
import { ISkill } from "./iSkill"; // Add the missing import statement
import { Spec } from "./spec";
import { getCopilotResponseAsString } from "../../../chat/utils";
import { ExecutionResultEnum } from "./executionResultEnum";

export class Explainer implements ISkill {
  name: string | undefined;
  capability: string | undefined;

  constructor() {
    this.name = "Explainer";
    this.capability = "Explain code snippet";
  }
  public canInvoke(spec: Spec): boolean {
    return (
      !!spec.userInput &&
      !!spec.appendix.codeSnippet &&
      !!spec.appendix.codeTaskBreakdown &&
      spec.appendix.codeTaskBreakdown.length > 0
    );
  }

  public async invoke(
    languageModel: LanguageModelChatMessage,
    response: ChatResponseStream,
    token: CancellationToken,
    spec: Spec
  ): Promise<{ result: ExecutionResultEnum; spec: Spec }> {
    const systemPrompt = `
Based on the user's input ${spec.userInput}, and the breakdown of the ask:
- ${spec.appendix.codeTaskBreakdown.join("\n- ")}

As output, you should only provide a very general short brief for the code snippet, not the code snippet itself. The output should be in the format of Markdown.
`;

    const userPrompt = `
Please explain the code snippet below:
\`\`\`typescript
${spec.appendix.codeSnippet}
\`\`\`

Let's think it step by step.
    `;

    // Perform the desired operation
    const messages: LanguageModelChatMessage[] = [
      new LanguageModelChatMessage(LanguageModelChatMessageRole.User, systemPrompt),
      new LanguageModelChatMessage(LanguageModelChatMessageRole.User, userPrompt),
    ];
    const copilotResponse = await getCopilotResponseAsString(
      "copilot-gpt-3.5-turbo",
      messages,
      token
    );
    spec.appendix.telemetryData.chatMessages.push(...messages);
    spec.appendix.telemetryData.responseChatMessages.push(
      new LanguageModelChatMessage(LanguageModelChatMessageRole.Assistant, copilotResponse)
    );

    if (!copilotResponse) {
      // something wrong with the LLM output
      // however it's not a hard stop, still ok produce the output without explanation
      return { result: ExecutionResultEnum.Failure, spec: spec };
    }

    spec.appendix.codeExplanation = copilotResponse;

    return { result: ExecutionResultEnum.Success, spec: spec };
  }
}

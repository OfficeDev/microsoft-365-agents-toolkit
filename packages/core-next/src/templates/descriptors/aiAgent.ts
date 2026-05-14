// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { TemplateDescriptor, TemplateActionOptions } from "../types";
import type { AtkContext } from "../../core/context";
import { scaffoldTemplates } from "../scaffold/scaffolder";
import { getTemplateReplaceMap } from "../scaffold/replaceMap";
import { type TemplateInfo, convertToLangKey } from "../scaffold/types";
import {
  llmProviderQuestion,
  azureOpenAiKeyQuestion,
  azureOpenAiEndpointQuestion,
  openAiKeyQuestion,
} from "../../questions/commonQuestions";

/**
 * AI Agent template artifact names matching the template repository folder names.
 */
export const AiAgentTemplateNames = {
  Chat: "custom-copilot-basic",
  RagAiSearch: "custom-copilot-rag-azure-ai-search",
  RagCustom: "custom-copilot-rag-customize",
} as const;

/**
 * Common LLM question specs reused across AI Agent descriptors.
 */
const llmQuestions = [
  llmProviderQuestion(),
  azureOpenAiKeyQuestion(),
  azureOpenAiEndpointQuestion(),
  openAiKeyQuestion(),
];

/**
 * Create a standard scaffold function for AI Agent templates.
 * AI Agents use the shared replaceMap which already includes LLM-related variables.
 */
function makeAiAgentScaffoldFn(templateName: string) {
  return async (ctx: AtkContext, opts: TemplateActionOptions) => {
    const replaceMap = getTemplateReplaceMap({
      appName: opts.projectName,
      ...opts,
    });

    const tplInfo: TemplateInfo = {
      templateName,
      language: convertToLangKey(opts.language),
      replaceMap,
    };

    const result = await scaffoldTemplates(ctx, [tplInfo], opts.destinationPath);
    return result.map((files) => ({
      projectPath: opts.destinationPath,
      warnings: files.length === 0 ? ["No files were scaffolded"] : undefined,
    }));
  };
}

/**
 * All AI Agent template descriptors.
 */
export const aiAgentTemplateDescriptors: TemplateDescriptor[] = [
  {
    id: "ai-agent/chat",
    name: "AI Chat Bot",
    description: "A Teams AI bot with conversational capabilities",
    category: "ai-agent",
    languages: ["typescript", "javascript", "csharp", "python"],
    scaffoldFn: makeAiAgentScaffoldFn(AiAgentTemplateNames.Chat),
    displayOrder: 1,
    questions: llmQuestions,
    tags: ["teamsApp", "publishable", "bot"],
  },
  {
    id: "ai-agent/rag-ai-search",
    name: "AI Bot with RAG (Azure AI Search)",
    description: "A Teams AI bot using Azure AI Search for retrieval-augmented generation",
    category: "ai-agent",
    languages: ["typescript", "javascript", "csharp", "python"],
    scaffoldFn: makeAiAgentScaffoldFn(AiAgentTemplateNames.RagAiSearch),
    displayOrder: 2,
    questions: llmQuestions,
    tags: ["teamsApp", "publishable", "bot"],
  },
  {
    id: "ai-agent/rag-custom",
    name: "AI Bot with RAG (Custom)",
    description: "A Teams AI bot with customizable RAG data source",
    category: "ai-agent",
    languages: ["typescript", "javascript", "csharp", "python"],
    scaffoldFn: makeAiAgentScaffoldFn(AiAgentTemplateNames.RagCustom),
    displayOrder: 3,
    questions: llmQuestions,
    tags: ["teamsApp", "publishable", "bot"],
  },
];

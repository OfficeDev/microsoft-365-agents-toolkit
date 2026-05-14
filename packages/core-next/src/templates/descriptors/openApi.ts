// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { TemplateDescriptor } from "../types";
import { makeOpenApiScaffoldFn } from "../openApi/scaffoldFn";
import {
  apiSpecPathQuestion,
  apiOperationsQuestion,
  llmProviderQuestion,
  azureOpenAiKeyQuestion,
  azureOpenAiEndpointQuestion,
  openAiKeyQuestion,
} from "../../questions/commonQuestions";

/**
 * OpenAPI-based template artifact names.
 */
export const OpenApiTemplateNames = {
  DaApiPlugin: "api-plugin-from-existing-api",
  AiAgentRagFromSpec: "custom-copilot-rag-custom-api",
  MeFromSpec: "copilot-plugin-existing-api",
} as const;

/**
 * OpenAPI-driven template descriptors.
 * These use the spec-parser adapter (currently stubbed) to generate code
 * from an OpenAPI specification after scaffolding a base template.
 */
export const openApiTemplateDescriptors: TemplateDescriptor[] = [
  {
    id: "da/api-plugin-from-spec",
    name: "Declarative Agent with API Plugin (from OpenAPI Spec)",
    description: "Generate an API plugin from an existing OpenAPI specification",
    category: "declarative-agent",
    languages: ["typescript", "javascript"],
    scaffoldFn: makeOpenApiScaffoldFn("Copilot", OpenApiTemplateNames.DaApiPlugin),
    displayOrder: 12,
    questions: [apiSpecPathQuestion(), apiOperationsQuestion()],
    tags: ["openapi"],
    testable: false,
  },
  {
    id: "ai-agent/rag-from-spec",
    name: "AI Bot with RAG (from OpenAPI Spec)",
    description: "Generate a Teams AI bot with RAG from an existing OpenAPI specification",
    category: "ai-agent",
    languages: ["typescript", "javascript", "csharp", "python"],
    scaffoldFn: makeOpenApiScaffoldFn("TeamsAi", OpenApiTemplateNames.AiAgentRagFromSpec),
    displayOrder: 6,
    questions: [
      apiSpecPathQuestion(),
      apiOperationsQuestion(),
      llmProviderQuestion(),
      azureOpenAiKeyQuestion(),
      azureOpenAiEndpointQuestion(),
      openAiKeyQuestion(),
    ],
    tags: ["openapi"],
    testable: false,
  },
  {
    id: "me/from-spec",
    name: "Message Extension (from OpenAPI Spec)",
    description: "Generate a message extension from an existing OpenAPI specification",
    category: "message-extension",
    languages: ["typescript", "javascript"],
    scaffoldFn: makeOpenApiScaffoldFn("SME", OpenApiTemplateNames.MeFromSpec),
    displayOrder: 7,
    questions: [apiSpecPathQuestion(), apiOperationsQuestion()],
    tags: ["openapi"],
    testable: false,
  },
];

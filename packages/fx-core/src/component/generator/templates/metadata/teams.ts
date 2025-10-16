// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateNames } from "../templateNames";
import { Template } from "./interface";

const chatWithYourDataTemplates: Template[] = [
  {
    id: "custom-copilot-rag-customize-ts",
    name: TemplateNames.CustomCopilotRagCustomize,
    language: "typescript",
    description: "",
  },
  {
    id: "custom-copilot-rag-customize-js",
    name: TemplateNames.CustomCopilotRagCustomize,
    language: "javascript",
    description: "",
  },
  {
    id: "custom-copilot-rag-customize-csharp",
    name: TemplateNames.CustomCopilotRagCustomize,
    language: "csharp",
    description: "",
  },
  {
    id: "custom-copilot-rag-azure-ai-search-ts",
    name: TemplateNames.CustomCopilotRagAzureAISearch,
    language: "typescript",
    description: "",
  },
  {
    id: "custom-copilot-rag-azure-ai-search-js",
    name: TemplateNames.CustomCopilotRagAzureAISearch,
    language: "javascript",
    description: "",
  },
  {
    id: "custom-copilot-rag-azure-ai-search-csharp",
    name: TemplateNames.CustomCopilotRagAzureAISearch,
    language: "csharp",
    description: "",
  },
  {
    id: "custom-copilot-rag-microsoft365-ts",
    name: TemplateNames.CustomCopilotRagMicrosoft365,
    language: "typescript",
    description: "",
  },
  {
    id: "custom-copilot-rag-microsoft365-js",
    name: TemplateNames.CustomCopilotRagMicrosoft365,
    language: "javascript",
    description: "",
  },
  {
    id: "custom-copilot-rag-microsoft365-csharp",
    name: TemplateNames.CustomCopilotRagMicrosoft365,
    language: "csharp",
    description: "",
  },
  {
    id: "custom-copilot-assistant-new-ts",
    name: TemplateNames.CustomCopilotAssistantNew,
    language: "typescript",
    description: "",
  },
  {
    id: "custom-copilot-assistant-new-js",
    name: TemplateNames.CustomCopilotAssistantNew,
    language: "javascript",
    description: "",
  },
  {
    id: "custom-copilot-assistant-new-csharp",
    name: TemplateNames.CustomCopilotAssistantNew,
    language: "csharp",
    description: "",
  },
  {
    id: "custom-copilot-assistant-assistants-api-ts",
    name: TemplateNames.CustomCopilotAssistantAssistantsApi,
    language: "typescript",
    description: "",
  },
  {
    id: "custom-copilot-assistant-assistants-api-js",
    name: TemplateNames.CustomCopilotAssistantAssistantsApi,
    language: "javascript",
    description: "",
  },
  {
    id: "custom-copilot-assistant-assistants-api-csharp",
    name: TemplateNames.CustomCopilotAssistantAssistantsApi,
    language: "csharp",
    description: "",
  },
];

const teamsOtherTemplates: Template[] = [
  {
    id: "basic-tab-ts",
    name: TemplateNames.Tab,
    language: "typescript",
    description: "Simple Teams Tab App",
  },
  {
    id: "default-bot-ts",
    name: TemplateNames.DefaultBot,
    language: "typescript",
    description: "",
  },
  {
    id: "default-bot-js",
    name: TemplateNames.DefaultBot,
    language: "javascript",
    description: "",
  },
  {
    id: "message-extension-v2-ts",
    name: TemplateNames.DefaultMessageExtension,
    language: "typescript",
    description: "",
  },
  // VS templates below
  {
    id: "basic-tab-csharp",
    name: TemplateNames.Tab,
    language: "csharp",
    description: "Simple Teams Tab App",
  },
  {
    id: "default-bot-csharp",
    name: TemplateNames.DefaultBot,
    language: "csharp",
    description: "",
  },
  {
    id: "notification-http-trigger-csharp",
    name: TemplateNames.NotificationHttpTrigger,
    language: "csharp",
    description: "",
  },
  {
    id: "notification-timer-trigger-csharp",
    name: TemplateNames.NotificationTimerTrigger,
    language: "csharp",
    description: "",
  },
  {
    id: "notification-http-timer-trigger-csharp",
    name: TemplateNames.NotificationHttpTimerTrigger,
    language: "csharp",
    description: "",
  },
  {
    id: "notification-webapi-csharp",
    name: TemplateNames.NotificationWebApi,
    language: "csharp",
    description: "",
  },
  {
    id: "command-and-response-csharp",
    name: TemplateNames.CommandAndResponse,
    language: "csharp",
    description: "",
  },
  {
    id: "workflow-csharp",
    name: TemplateNames.Workflow,
    language: "csharp",
    description: "",
  },
  {
    id: "message-extension-v2-csharp",
    name: TemplateNames.DefaultMessageExtension,
    language: "csharp",
    description: "",
  },
];

export const teamsAgentsAndAppsTemplates: Template[] = [
  {
    id: "custom-copilot-basic-ts",
    name: TemplateNames.CustomCopilotBasic,
    language: "typescript",
    description: "",
  },
  {
    id: "custom-copilot-basic-js",
    name: TemplateNames.CustomCopilotBasic,
    language: "javascript",
    description: "",
  },
  {
    id: "custom-copilot-basic-csharp",
    name: TemplateNames.CustomCopilotBasic,
    language: "csharp",
    description: "",
  },
  ...chatWithYourDataTemplates,
  {
    id: "teams-collaborator-agent-ts",
    name: TemplateNames.TeamsCollaboratorAgent,
    language: "typescript",
    description: "",
  },
  ...teamsOtherTemplates,
];

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateNames } from "../templateNames";
import { Template } from "./interface";

const chatWithYourDataTemplates: Template[] = [
  {
    id: "custom-copilot-rag-customize-ts",
    name: TemplateNames.CustomCopilotRagCustomize,
    language: "typescript",
    description:
      "Build a custom AI agent that can chat with your data using RAG (Retrieval Augmented Generation)",
  },
  {
    id: "custom-copilot-rag-customize-js",
    name: TemplateNames.CustomCopilotRagCustomize,
    language: "javascript",
    description:
      "Build a custom AI agent that can chat with your data using RAG (Retrieval Augmented Generation)",
  },
  {
    id: "custom-copilot-rag-customize-csharp",
    name: TemplateNames.CustomCopilotRagCustomize,
    language: "csharp",
    description:
      "Build a custom AI agent that can chat with your data using RAG (Retrieval Augmented Generation)",
  },
  {
    id: "custom-copilot-rag-customize-python",
    name: TemplateNames.CustomCopilotRagCustomize,
    language: "python",
    description:
      "Build a custom AI agent that can chat with your data using RAG (Retrieval Augmented Generation)",
  },
  {
    id: "custom-copilot-rag-azure-ai-search-ts",
    name: TemplateNames.CustomCopilotRagAzureAISearch,
    language: "typescript",
    description: "AI agent with Azure AI Search integration for intelligent data retrieval",
  },
  {
    id: "custom-copilot-rag-azure-ai-search-js",
    name: TemplateNames.CustomCopilotRagAzureAISearch,
    language: "javascript",
    description: "AI agent with Azure AI Search integration for intelligent data retrieval",
  },
  {
    id: "custom-copilot-rag-azure-ai-search-csharp",
    name: TemplateNames.CustomCopilotRagAzureAISearch,
    language: "csharp",
    description: "AI agent with Azure AI Search integration for intelligent data retrieval",
  },
  {
    id: "custom-copilot-rag-azure-ai-search-python",
    name: TemplateNames.CustomCopilotRagAzureAISearch,
    language: "python",
    description: "AI agent with Azure AI Search integration for intelligent data retrieval",
  },
  {
    id: "custom-copilot-rag-microsoft365-ts",
    name: TemplateNames.CustomCopilotRagMicrosoft365,
    language: "typescript",
    description: "AI agent that searches and reasons over Microsoft 365 content",
  },
  {
    id: "custom-copilot-rag-microsoft365-js",
    name: TemplateNames.CustomCopilotRagMicrosoft365,
    language: "javascript",
    description: "AI agent that searches and reasons over Microsoft 365 content",
  },
  {
    id: "custom-copilot-rag-microsoft365-csharp",
    name: TemplateNames.CustomCopilotRagMicrosoft365,
    language: "csharp",
    description: "AI agent that searches and reasons over Microsoft 365 content",
  },
  {
    id: "custom-copilot-assistant-new-ts",
    name: TemplateNames.CustomCopilotAssistantNew,
    language: "typescript",
    description: "Create a new AI assistant with custom capabilities",
  },
  {
    id: "custom-copilot-assistant-new-js",
    name: TemplateNames.CustomCopilotAssistantNew,
    language: "javascript",
    description: "Create a new AI assistant with custom capabilities",
  },
  {
    id: "custom-copilot-assistant-new-csharp",
    name: TemplateNames.CustomCopilotAssistantNew,
    language: "csharp",
    description: "Create a new AI assistant with custom capabilities",
  },
  {
    id: "custom-copilot-assistant-assistants-api-ts",
    name: TemplateNames.CustomCopilotAssistantAssistantsApi,
    language: "typescript",
    description: "Build an AI assistant using OpenAI Assistants API",
  },
  {
    id: "custom-copilot-assistant-assistants-api-js",
    name: TemplateNames.CustomCopilotAssistantAssistantsApi,
    language: "javascript",
    description: "Build an AI assistant using OpenAI Assistants API",
  },
  {
    id: "custom-copilot-assistant-assistants-api-csharp",
    name: TemplateNames.CustomCopilotAssistantAssistantsApi,
    language: "csharp",
    description: "Build an AI assistant using OpenAI Assistants API",
  },
];

const teamsOtherTemplates: Template[] = [
  {
    id: "basic-tab-ts",
    name: TemplateNames.Tab,
    language: "typescript",
    description: "A simple implementation of a web app that's ready to customize",
  },
  {
    id: "default-bot-ts",
    name: TemplateNames.DefaultBot,
    language: "typescript",
    description: "A simple implementation of an echo bot that's ready for customization",
  },
  {
    id: "default-bot-js",
    name: TemplateNames.DefaultBot,
    language: "javascript",
    description: "A simple implementation of an echo bot that's ready for customization",
  },
  {
    id: "default-bot-python",
    name: TemplateNames.DefaultBot,
    language: "python",
    description: "A simple implementation of an echo bot that's ready for customization",
  },
  {
    id: "message-extension-v2-ts",
    name: TemplateNames.DefaultMessageExtension,
    language: "typescript",
    description: "Receive user input, process it, and send customized results",
  },
  {
    id: "message-extension-v2-python",
    name: TemplateNames.DefaultMessageExtension,
    language: "python",
    description: "Receive user input, process it, and send customized results",
  },
];

export const teamsAgentsAndAppsTemplates: Template[] = [
  {
    id: "custom-copilot-basic-ts",
    name: TemplateNames.CustomCopilotBasic,
    language: "typescript",
    description: "Build a basic AI chatbot for Microsoft Teams",
  },
  {
    id: "custom-copilot-basic-js",
    name: TemplateNames.CustomCopilotBasic,
    language: "javascript",
    description: "Build a basic AI chatbot for Microsoft Teams",
  },
  {
    id: "custom-copilot-basic-csharp",
    name: TemplateNames.CustomCopilotBasic,
    language: "csharp",
    description: "Build a basic AI chatbot for Microsoft Teams",
  },
  {
    id: "custom-copilot-basic-python",
    name: TemplateNames.CustomCopilotBasic,
    language: "python",
    description: "Build a basic AI chatbot for Microsoft Teams",
  },
  ...chatWithYourDataTemplates,
  {
    id: "teams-collaborator-agent-ts",
    name: TemplateNames.TeamsCollaboratorAgent,
    language: "typescript",
    description: "AI agent that can collaborate with users in Microsoft Teams",
  },
  ...teamsOtherTemplates,
];

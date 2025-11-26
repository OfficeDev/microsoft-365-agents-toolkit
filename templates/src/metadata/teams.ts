// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateNames } from "../templateNames";
import { getString } from "../ui/helper";
import { Template } from "./interface";

const chatWithYourDataTemplates: Template[] = [
  {
    id: "custom-copilot-rag-customize-ts",
    name: TemplateNames.CustomCopilotRagCustomize,
    language: "typescript",
    displayName: "Teams Agent with Data from Customized Source",
    description: getString("template.teams.rag.detail"),
  },
  {
    id: "custom-copilot-rag-customize-js",
    name: TemplateNames.CustomCopilotRagCustomize,
    language: "javascript",
    displayName: "Teams Agent with Data from Customized Source",
    description: getString("template.teams.rag.detail"),
  },
  {
    id: "custom-copilot-rag-customize-csharp",
    name: TemplateNames.CustomCopilotRagCustomize,
    language: "csharp",
    displayName: "Teams Agent with Data from Customized Source",
    description: getString("template.teams.rag.detail"),
  },
  {
    id: "custom-copilot-rag-customize-python",
    name: TemplateNames.CustomCopilotRagCustomize,
    language: "python",
    displayName: "Teams Agent with Data from Customized Source",
    description: getString("template.teams.rag.detail"),
  },
  {
    id: "custom-copilot-rag-azure-ai-search-ts",
    name: TemplateNames.CustomCopilotRagAzureAISearch,
    language: "typescript",
    displayName: "Teams Agent with Data from Azure AI Search",
    description: getString("template.teams.rag.detail"),
  },
  {
    id: "custom-copilot-rag-azure-ai-search-js",
    name: TemplateNames.CustomCopilotRagAzureAISearch,
    language: "javascript",
    displayName: "Teams Agent with Data from Azure AI Search",
    description: getString("template.teams.rag.detail"),
  },
  {
    id: "custom-copilot-rag-azure-ai-search-csharp",
    name: TemplateNames.CustomCopilotRagAzureAISearch,
    language: "csharp",
    displayName: "Teams Agent with Data from Azure AI Search",
    description: getString("template.teams.rag.detail"),
  },
  {
    id: "custom-copilot-rag-azure-ai-search-python",
    name: TemplateNames.CustomCopilotRagAzureAISearch,
    language: "python",
    displayName: "Teams Agent with Data from Azure AI Search",
    description: getString("template.teams.rag.detail"),
  },
  // {
  //   id: "custom-copilot-rag-microsoft365-ts",
  //   name: TemplateNames.CustomCopilotRagMicrosoft365,
  //   language: "typescript",
  //   displayName: "AI Agent with Microsoft 365 (RAG)",
  //   description: "AI agent that searches and reasons over Microsoft 365 content",
  // },
  // {
  //   id: "custom-copilot-rag-microsoft365-js",
  //   name: TemplateNames.CustomCopilotRagMicrosoft365,
  //   language: "javascript",
  //   displayName: "AI Agent with Microsoft 365 (RAG)",
  //   description: "AI agent that searches and reasons over Microsoft 365 content",
  // },
  // {
  //   id: "custom-copilot-rag-microsoft365-csharp",
  //   name: TemplateNames.CustomCopilotRagMicrosoft365,
  //   language: "csharp",
  //   displayName: "AI Agent with Microsoft 365 (RAG)",
  //   description: "AI agent that searches and reasons over Microsoft 365 content",
  // },
  // {
  //   id: "custom-copilot-assistant-new-ts",
  //   name: TemplateNames.CustomCopilotAssistantNew,
  //   language: "typescript",
  //   displayName: "AI Assistant (New)",
  //   description: "Create a new AI assistant with custom capabilities",
  // },
  // {
  //   id: "custom-copilot-assistant-new-js",
  //   name: TemplateNames.CustomCopilotAssistantNew,
  //   language: "javascript",
  //   displayName: "AI Assistant (New)",
  //   description: "Create a new AI assistant with custom capabilities",
  // },
  // {
  //   id: "custom-copilot-assistant-new-csharp",
  //   name: TemplateNames.CustomCopilotAssistantNew,
  //   language: "csharp",
  //   displayName: "AI Assistant (New)",
  //   description: "Create a new AI assistant with custom capabilities",
  // },
  // {
  //   id: "custom-copilot-assistant-assistants-api-ts",
  //   name: TemplateNames.CustomCopilotAssistantAssistantsApi,
  //   language: "typescript",
  //   displayName: "AI Assistant (Assistants API)",
  //   description: "Build an AI assistant using OpenAI Assistants API",
  // },
  // {
  //   id: "custom-copilot-assistant-assistants-api-js",
  //   name: TemplateNames.CustomCopilotAssistantAssistantsApi,
  //   language: "javascript",
  //   displayName: "AI Assistant (Assistants API)",
  //   description: "Build an AI assistant using OpenAI Assistants API",
  // },
  // {
  //   id: "custom-copilot-assistant-assistants-api-csharp",
  //   name: TemplateNames.CustomCopilotAssistantAssistantsApi,
  //   language: "csharp",
  //   displayName: "AI Assistant (Assistants API)",
  //   description: "Build an AI assistant using OpenAI Assistants API",
  // },
];

const teamsOtherTemplates: Template[] = [
  {
    id: "basic-tab-ts",
    name: TemplateNames.Tab,
    language: "typescript",
    displayName: getString("template.teams.others.tab.label"),
    description: getString("template.teams.others.tab.detail"),
  },
  {
    id: "default-bot-ts",
    name: TemplateNames.DefaultBot,
    language: "typescript",
    displayName: getString("template.teams.others.bot.label"),
    description: getString("template.teams.others.bot.detail"),
  },
  {
    id: "default-bot-js",
    name: TemplateNames.DefaultBot,
    language: "javascript",
    displayName: getString("template.teams.others.bot.label"),
    description: getString("template.teams.others.bot.detail"),
  },
  {
    id: "default-bot-python",
    name: TemplateNames.DefaultBot,
    language: "python",
    displayName: getString("template.teams.others.bot.label"),
    description: getString("template.teams.others.bot.detail"),
  },
  {
    id: "message-extension-v2-ts",
    name: TemplateNames.DefaultMessageExtension,
    language: "typescript",
    displayName: getString("template.teams.others.messageExtension.label"),
    description: getString("template.teams.others.messageExtension.detail"),
  },
  {
    id: "message-extension-v2-python",
    name: TemplateNames.DefaultMessageExtension,
    language: "python",
    displayName: getString("template.teams.others.messageExtension.label"),
    description: getString("template.teams.others.messageExtension.detail"),
  },
];

export const teamsAgentsAndAppsTemplates: Template[] = [
  {
    id: "custom-copilot-basic-ts",
    name: TemplateNames.CustomCopilotBasic,
    language: "typescript",
    displayName: getString("template.teams.general.label"),
    description: getString("template.teams.general.detail"),
  },
  {
    id: "custom-copilot-basic-js",
    name: TemplateNames.CustomCopilotBasic,
    language: "javascript",
    displayName: getString("template.teams.general.label"),
    description: getString("template.teams.general.detail"),
  },
  {
    id: "custom-copilot-basic-csharp",
    name: TemplateNames.CustomCopilotBasic,
    language: "csharp",
    displayName: getString("template.teams.general.label"),
    description: getString("template.teams.general.detail"),
  },
  {
    id: "custom-copilot-basic-python",
    name: TemplateNames.CustomCopilotBasic,
    language: "python",
    displayName: getString("template.teams.general.label"),
    description: getString("template.teams.general.detail"),
  },
  ...chatWithYourDataTemplates,
  {
    id: "teams-collaborator-agent-ts",
    name: TemplateNames.TeamsCollaboratorAgent,
    language: "typescript",
    displayName: getString("template.teams.collaboratorAgent.label"),
    description: getString("template.teams.collaboratorAgent.detail"),
  },
  {
    id: "teams-collaborator-agent-csharp",
    name: TemplateNames.TeamsCollaboratorAgent,
    language: "csharp",
    displayName: getString("template.teams.collaboratorAgent.label"),
    description: getString("template.teams.collaboratorAgent.detail"),
  },
  ...teamsOtherTemplates,
];

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface Template {
  id: string;
  name: string;
  language: "typescript" | "javascript" | "csharp" | "python" | "none";
  description: string;
  link?: string;
}

export enum TemplateNames {
  Empty = "empty",
  Tab = "non-sso-tab",
  SsoTab = "sso-tab",
  SsoTabObo = "sso-tab-with-obo-flow",
  TabSSR = "non-sso-tab-ssr",
  SsoTabSSR = "sso-tab-ssr",
  DashboardTab = "dashboard-tab",
  NotificationExpress = "notification-express",
  NotificationWebApi = "notification-webapi",
  NotificationHttpTriggerIsolated = "notification-http-trigger-isolated",
  NotificationHttpTrigger = "notification-http-trigger",
  NotificationTimerTriggerIsolated = "notification-timer-trigger-isolated",
  NotificationTimerTrigger = "notification-timer-trigger",
  NotificationHttpTimerTriggerIsolated = "notification-http-timer-trigger-isolated",
  NotificationHttpTimerTrigger = "notification-http-timer-trigger",
  CommandAndResponse = "command-and-response",
  Workflow = "workflow",
  DefaultBot = "default-bot",
  MessageExtension = "message-extension",
  MessageExtensionAction = "message-extension-action",
  MessageExtensionSearch = "message-extension-search",
  MessageExtensionCopilot = "message-extension-copilot",
  M365MessageExtension = "m365-message-extension",
  TabAndDefaultBot = "non-sso-tab-default-bot",
  BotAndMessageExtension = "default-bot-message-extension",
  LinkUnfurling = "link-unfurling",
  AIBot = "ai-bot",
  AIAssistantBot = "ai-assistant-bot",
  ApiPluginFromScratch = "api-plugin-from-scratch",
  ApiPluginFromScratchBearer = "api-plugin-from-scratch-bearer",
  ApiPluginFromScratchOAuth = "api-plugin-from-scratch-oauth",
  CopilotPluginFromScratch = "copilot-plugin-from-scratch",
  CopilotPluginFromScratchApiKey = "copilot-plugin-from-scratch-api-key",
  ApiMessageExtensionSso = "api-message-extension-sso",
  CustomCopilotBasic = "custom-copilot-basic",
  CustomCopilotRagCustomize = "custom-copilot-rag-customize",
  CustomCopilotRagAzureAISearch = "custom-copilot-rag-azure-ai-search",
  CustomCopilotRagCustomApi = "custom-copilot-rag-custom-api",
  CustomCopilotRagMicrosoft365 = "custom-copilot-rag-microsoft365",
  CustomCopilotAssistantNew = "custom-copilot-assistant-new",
  CustomCopilotAssistantAssistantsApi = "custom-copilot-assistant-assistants-api",
  BasicGpt = "copilot-gpt-basic",
  GptWithPluginFromScratch = "copilot-gpt-from-scratch-plugin",
  DeclarativeAgentWithApiSpec = "declarative-agent-with-api-spec",
  OutlookTaskpane = "office-addin-outlook-taskpane",
  WXPTaskpane = "office-addin-wxp-taskpane",
  OfficeAddinCommon = "office-addin-common",
}

export const Templates: Template[] = [
  {
    id: "notification-express-ts",
    name: TemplateNames.NotificationExpress,
    language: "typescript",
    description: "Notification Express",
  },
  {
    id: "api-plugin-existing-api",
    name: "api-plugin-existing-api",
    language: "none",
    description: "Declarative agent with api spec",
  },
];

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface Template {
  id: string;
  name: string;
  language: "typescript" | "javascript" | "csharp" | "python" | "none" | "common";
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
  MessageExtensionM365 = "m365-message-extension",
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
  OfficeAddinCommon = "office-addin-config",
  ApiPluginExistingApi = "api-plugin-existing-api",
  CopilotPluginExistingApi = "copilot-plugin-existing-api",
}

const tabTemplates: Template[] = [
  {
    id: "non-sso-tab-ts",
    name: TemplateNames.Tab,
    language: "typescript",
    description: "Simple Teams Tab App",
  },
  {
    id: "non-sso-tab-js",
    name: TemplateNames.Tab,
    language: "javascript",
    description: "Simple Teams Tab App",
  },
  {
    id: "non-sso-tab-ssr-cs",
    name: TemplateNames.TabSSR,
    language: "csharp",
    description: "Simple Teams Tab App",
  },
  {
    id: "sso-tab-ssr-cs",
    name: TemplateNames.SsoTabSSR,
    language: "csharp",
    description: "Simple Teams Tab App with SSO",
  },
  {
    id: "sso-tab-with-obo-flow-ts",
    name: TemplateNames.SsoTabObo,
    language: "typescript",
    description: "Simple Teams Tab App with OBO Flow",
  },
  {
    id: "sso-tab-with-obo-flow-js",
    name: TemplateNames.SsoTabObo,
    language: "javascript",
    description: "Simple Teams Tab App with OBO Flow",
  },
  {
    id: "dashboard-tab-ts",
    name: TemplateNames.DashboardTab,
    language: "typescript",
    description: "Dashboard Tab App",
    link: "https://aka.ms/teamsfx-dashboard-app",
  },
  {
    id: "dashboard-tab-js",
    name: TemplateNames.DashboardTab,
    language: "javascript",
    description: "Dashboard Tab App",
    link: "https://aka.ms/teamsfx-dashboard-app",
  },
];

const basicBotTemplates: Template[] = [
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
    id: "workflow-ts",
    name: TemplateNames.Workflow,
    language: "typescript",
    description: "",
    link: "https://aka.ms/teamsfx-workflow",
  },
  {
    id: "workflow-js",
    name: TemplateNames.Workflow,
    language: "javascript",
    description: "",
    link: "https://aka.ms/teamsfx-workflow",
  },
  {
    id: "command-and-response-ts",
    name: TemplateNames.CommandAndResponse,
    language: "typescript",
    description: "",
    link: "https://aka.ms/teamsfx-create-command",
  },
  {
    id: "command-and-response-js",
    name: TemplateNames.CommandAndResponse,
    language: "javascript",
    description: "",
    link: "https://aka.ms/teamsfx-create-command",
  },
];

const notificationBotTemplates: Template[] = [
  {
    id: "notification-express-ts",
    name: TemplateNames.NotificationExpress,
    language: "typescript",
    description: "",
  },
  {
    id: "notification-express-js",
    name: TemplateNames.NotificationExpress,
    language: "javascript",
    description: "",
  },
  {
    id: "notification-http-trigger-ts",
    name: TemplateNames.NotificationHttpTrigger,
    language: "typescript",
    description: "",
  },
  {
    id: "notification-http-trigger-js",
    name: TemplateNames.NotificationHttpTrigger,
    language: "javascript",
    description: "",
  },
  {
    id: "notification-timer-trigger-ts",
    name: TemplateNames.NotificationTimerTrigger,
    language: "typescript",
    description: "",
  },
  {
    id: "notification-timer-trigger-js",
    name: TemplateNames.NotificationTimerTrigger,
    language: "javascript",
    description: "",
  },
  {
    id: "notification-http-timer-trigger-ts",
    name: TemplateNames.NotificationHttpTimerTrigger,
    language: "typescript",
    description: "",
  },
  {
    id: "notification-http-timer-trigger-js",
    name: TemplateNames.NotificationHttpTimerTrigger,
    language: "javascript",
    description: "",
  },
];

const messageExtensionTemplates: Template[] = [
  {
    id: "message-extension-ts",
    name: TemplateNames.MessageExtension,
    language: "typescript",
    description: "",
  },
  {
    id: "message-extension-js",
    name: TemplateNames.MessageExtension,
    language: "javascript",
    description: "",
  },
  {
    id: "m365-message-extension-ts",
    name: TemplateNames.MessageExtensionM365,
    language: "typescript",
    description: "",
  },
  {
    id: "m365-message-extension-js",
    name: TemplateNames.MessageExtensionM365,
    language: "javascript",
    description: "",
  },
  {
    id: "message-extension-action-ts",
    name: TemplateNames.MessageExtensionAction,
    language: "typescript",
    description: "",
  },
  {
    id: "message-extension-action-js",
    name: TemplateNames.MessageExtensionAction,
    language: "javascript",
    description: "",
  },
  {
    id: "message-extension-copilot-ts",
    name: TemplateNames.MessageExtensionCopilot,
    language: "typescript",
    description: "",
  },
  {
    id: "message-extension-copilot-js",
    name: TemplateNames.MessageExtensionCopilot,
    language: "javascript",
    description: "",
  },
  {
    id: "link-unfurling-ts",
    name: TemplateNames.LinkUnfurling,
    language: "typescript",
    description: "",
  },
  {
    id: "link-unfurling-js",
    name: TemplateNames.LinkUnfurling,
    language: "javascript",
    description: "",
  },
  {
    id: "api-plugin-existing-api",
    name: TemplateNames.ApiPluginExistingApi,
    language: "common",
    description: "",
  },
];

const copilotPluginTemplates: Template[] = [
  {
    id: "copilot-plugin-from-scratch-ts",
    name: TemplateNames.CopilotPluginFromScratch,
    language: "typescript",
    description: "",
  },
  {
    id: "copilot-plugin-from-scratch-js",
    name: TemplateNames.CopilotPluginFromScratch,
    language: "javascript",
    description: "",
  },
  {
    id: "copilot-plugin-from-scratch-api-key-ts",
    name: TemplateNames.CopilotPluginFromScratchApiKey,
    language: "typescript",
    description: "",
  },
  {
    id: "copilot-plugin-from-scratch-api-key-js",
    name: TemplateNames.CopilotPluginFromScratchApiKey,
    language: "javascript",
    description: "",
  },
];

export const customEngineAgentTemplates: Template[] = [
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
    id: "custom-copilot-assistant-assistants-apits",
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
];

const tdpTemplates: Template[] = [
  {
    id: "non-sso-tab-default-bot-ts",
    name: TemplateNames.TabAndDefaultBot,
    language: "typescript",
    description: "",
  },
  {
    id: "non-sso-tab-default-bot-js",
    name: TemplateNames.TabAndDefaultBot,
    language: "javascript",
    description: "",
  },
  {
    id: "default-bot-message-extension-ts",
    name: TemplateNames.BotAndMessageExtension,
    language: "typescript",
    description: "",
  },
  {
    id: "default-bot-message-extension-js",
    name: TemplateNames.BotAndMessageExtension,
    language: "javascript",
    description: "",
  },
];

export const Templates: Template[] = [
  ...tabTemplates,
  ...basicBotTemplates,
  ...notificationBotTemplates,
  ...messageExtensionTemplates,
  ...copilotPluginTemplates,
  ...customEngineAgentTemplates,
];

export const TdpTemplates: Template[] = [...tdpTemplates];

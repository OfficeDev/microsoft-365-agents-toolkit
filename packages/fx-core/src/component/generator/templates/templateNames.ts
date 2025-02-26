// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// Sorted templates that maps to question tree
// @author Ning Tang
export enum TemplateNames {
  // declarative agent
  BasicGpt = "copilot-gpt-basic", // TODO: handled by xxx generator
  ApiPluginFromScratch = "api-plugin-from-scratch", // TODO: handled by xxx generator
  ApiPluginFromScratchBearer = "api-plugin-from-scratch-bearer", // TODO: handled by xxx generator (The ApiPluginFromScratchBearer template is currently actually ApiPluginFromScratchAPIKey)
  ApiPluginFromScratchOAuth = "api-plugin-from-scratch-oauth", // TODO: handled by xxx generator
  DeclarativeAgentWithApiSpec = "declarative-agent-with-api-spec", // TODO: split to multiple templates, it's mapped in multiple options now.
  ApiPluginWithExistingApiSpec = "api-plugin-existing-api",

  // custom engine agent
  CustomCopilotBasic = "custom-copilot-basic",
  CustomCopilotRagCustomize = "custom-copilot-rag-customize",
  CustomCopilotRagAzureAISearch = "custom-copilot-rag-azure-ai-search",
  CustomCopilotRagCustomApi = "custom-copilot-rag-custom-api", // TODO: handled by xxx generator
  CustomCopilotRagMicrosoft365 = "custom-copilot-rag-microsoft365",
  CustomCopilotAssistantNew = "custom-copilot-assistant-new",
  CustomCopilotAssistantAssistantsApi = "custom-copilot-assistant-assistants-api",

  // tab
  Tab = "non-sso-tab",
  SsoTabObo = "sso-tab-with-obo-flow",
  DashboardTab = "dashboard-tab",
  TabSSR = "non-sso-tab-ssr", // handled by SsrTabGenerator
  SsoTabSSR = "sso-tab-ssr", // handled by SsrTabGenerator

  // bot
  DefaultBot = "default-bot",
  NotificationExpress = "notification-express", // vsc only
  NotificationWebApi = "notification-webapi", // vs only
  NotificationHttpTrigger = "notification-http-trigger",
  NotificationTimerTrigger = "notification-timer-trigger",
  NotificationHttpTimerTrigger = "notification-http-timer-trigger",
  CommandAndResponse = "command-and-response",
  Workflow = "workflow",

  // messaging extension
  MessageExtensionWithNewApiFromScratch = "copilot-plugin-from-scratch",
  MessageExtensionWithNewApiFromScratchUsingApiKey = "copilot-plugin-from-scratch-api-key",
  MessageExtensionWithNewApiFromScratchUsingOAuth = "api-message-extension-sso",
  MessageExtensionWithExistingApiSpec = "copilot-plugin-existing-api", // TODO: handled by xxx generator
  MessageExtensionM365 = "m365-message-extension",
  MessageExtensionAction = "message-extension-action",
  LinkUnfurling = "link-unfurling",

  // WXP
  OutlookTaskpane = "office-addin-outlook-taskpane", // handled by OfficeAddinGeneratorNew
  WXPTaskpane = "office-addin-wxpo-taskpane", // handled by OfficeAddinGeneratorNew
  OfficeAddinCommon = "office-addin-config", // handled by OfficeAddinGeneratorNew

  // from TDP only
  TabAndDefaultBot = "non-sso-tab-default-bot",
  BotAndMessageExtension = "default-bot-message-extension",
  MessageExtension = "message-extension",

  // VS only
  Empty = "empty",
  AIBot = "ai-bot",
  AIAssistantBot = "ai-assistant-bot",
  MessageExtensionSearch = "message-extension-search",

  // not used yet
}

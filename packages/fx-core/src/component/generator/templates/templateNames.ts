// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// Sorted templates that maps to question tree
// @author Ning Tang
export enum TemplateNames {
  // declarative agent
  DeclarativeAgentBasic = "declarative-agent-basic", // handled by DeclarativeAgentGenerator
  DeclarativeAgentWithActionFromScratch = "declarative-agent-with-action-from-scratch", // handled by DeclarativeAgentGenerator
  DeclarativeAgentWithActionFromScratchBearer = "declarative-agent-with-action-from-scratch-bearer", // handled by DeclarativeAgentGenerator (The DeclarativeAgentWithActionFromScratchBearer template is currently actually ApiPluginFromScratchAPIKey)
  DeclarativeAgentWithActionFromScratchOAuth = "declarative-agent-with-action-from-scratch-oauth", // handled by DeclarativeAgentGenerator
  DeclarativeAgentWithActionFromExistingApiSpec = "declarative-agent-with-action-from-existing-api", // handled by DeclarativeAgentWithExistingApiSpecGenerator
  DeclarativeAgentWithExistingAction = "declarative-agent-with-existing-action", // handled by DeclarativeAgentGenerator

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
  MessageExtensionWithNewApiFromScratch = "message-extension-with-api-from-scratch",
  MessageExtensionWithNewApiFromScratchUsingApiKey = "message-extension-with-api-from-scratch-api-key",
  MessageExtensionWithNewApiFromScratchUsingOAuth = "message-extension-with-api-from-scratch-sso",
  MessageExtensionWithExistingApiSpec = "message-extension-with-existing-api", // handled by MessageExtensionWithExistingApiSpecGenerator
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
}

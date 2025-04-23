// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// Sorted templates that maps to question tree
// @author Ning Tang
export enum TemplateNames {
  // declarative agent
  DeclarativeAgentBasic = "copilot-gpt-basic", // handled by DeclarativeAgentGenerator
  DeclarativeAgentWithActionFromScratch = "api-plugin-from-scratch", // handled by DeclarativeAgentGenerator
  DeclarativeAgentWithActionFromScratchBearer = "api-plugin-from-scratch-bearer", // handled by DeclarativeAgentGenerator (The DeclarativeAgentWithActionFromScratchBearer template is currently actually ApiPluginFromScratchAPIKey)
  DeclarativeAgentWithActionFromScratchOAuth = "api-plugin-from-scratch-oauth", // handled by DeclarativeAgentGenerator
  DeclarativeAgentWithActionFromExistingApiSpec = "api-plugin-from-existing-api", // handled by DeclarativeAgentWithExistingApiSpecGenerator
  DeclarativeAgentWithExistingAction = "api-plugin-existing-api", // handled by DeclarativeAgentGenerator
  DeclarativeAgentWithTypeSpec = "declarative-agent-typespec", // handled by DeclarativeAgentGenerator
  DeclarativeAgentWithGraphConnector = "declarative-agent-with-graph-connector", // handled by DeclarativeAgentGenerator

  // custom engine agent
  BasicCustomEngineAgent = "basic-custom-engine-agent",
  CustomCopilotWeatherAgent = "custom-copilot-weather-agent", // only supported in VS, using agent SDK

  // agent for Teams
  CustomCopilotBasic = "custom-copilot-basic",
  CustomCopilotRagCustomize = "custom-copilot-rag-customize",
  CustomCopilotRagAzureAISearch = "custom-copilot-rag-azure-ai-search",
  CustomCopilotRagCustomApi = "custom-copilot-rag-custom-api", // handled by CustomEngineAgentWithExistingApiSpecGenerator
  CustomCopilotRagMicrosoft365 = "custom-copilot-rag-microsoft365",
  CustomCopilotAssistantNew = "custom-copilot-assistant-new",
  CustomCopilotAssistantAssistantsApi = "custom-copilot-assistant-assistants-api",

  // copilot connector
  GraphConnector = "graph-connector", // vsc only

  // tab
  Tab = "non-sso-tab",
  SsoTabObo = "sso-tab-with-obo-flow",
  DashboardTab = "dashboard-tab",
  TabSSR = "non-sso-tab-ssr", // handled by SsrTabGenerator
  SsoTabSSR = "sso-tab-ssr", // handled by SsrTabGenerator
  TabSPFx = "spfx-tab", // handled by SPFxGeneratorNew

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
  MessageExtensionWithExistingApiSpec = "copilot-plugin-existing-api", // handled by MessageExtensionWithExistingApiSpecGenerator
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

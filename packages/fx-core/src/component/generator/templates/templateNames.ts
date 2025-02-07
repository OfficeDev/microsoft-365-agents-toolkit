// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { Inputs } from "@microsoft/teamsfx-api";
import {
  ApiAuthOptions,
  CapabilityOptions,
  CustomCopilotAssistantOptions,
  CustomCopilotRagOptions,
  MeArchitectureOptions,
  NotificationTriggerOptions,
  ProgrammingLanguage,
  QuestionNames,
} from "../../../question/constants";

export enum TemplateNames {
  // tab
  Tab = "non-sso-tab",
  SsoTabObo = "sso-tab-with-obo-flow",
  DashboardTab = "dashboard-tab",
  TabSSR = "non-sso-tab-ssr", // handled by SsrTabGenerator
  SsoTabSSR = "sso-tab-ssr", // handled by SsrTabGenerator

  // bot
  DefaultBot = "default-bot",
  Workflow = "workflow",
  CommandAndResponse = "command-and-response",
  NotificationHttpTrigger = "notification-http-trigger",
  NotificationTimerTrigger = "notification-timer-trigger",
  NotificationHttpTimerTrigger = "notification-http-timer-trigger",
  // bot (platform specific)
  NotificationExpress = "notification-express",
  NotificationWebApi = "notification-webapi",

  // messaging extension
  MessageExtension = "message-extension",
  MessageExtensionAction = "message-extension-action",
  LinkUnfurling = "link-unfurling",
  ApiMessageExtensionSso = "api-message-extension-sso",
  // messaging extension (platform specific)
  ApiPluginExistingApi = "api-plugin-existing-api", // TODO: handled by xxx generator
  MessageExtensionM365 = "m365-message-extension",
  MessageExtensionCopilot = "message-extension-copilot",
  MessageExtensionSearch = "message-extension-search",

  // copilot plugin
  CopilotPluginFromScratch = "copilot-plugin-from-scratch",
  CopilotPluginFromScratchApiKey = "copilot-plugin-from-scratch-api-key",
  // copilot plugin (platform specific)
  CopilotPluginExistingApi = "copilot-plugin-existing-api",

  // declarative agent
  BasicGpt = "copilot-gpt-basic", // TODO: handled by xxx generator
  ApiPluginFromScratch = "api-plugin-from-scratch", // TODO: handled by xxx generator
  ApiPluginFromScratchBearer = "api-plugin-from-scratch-bearer", // TODO: handled by xxx generator
  ApiPluginFromScratchOAuth = "api-plugin-from-scratch-oauth", // TODO: handled by xxx generator
  DeclarativeAgentWithApiSpec = "declarative-agent-with-api-spec",

  // custom engine agent
  CustomCopilotBasic = "custom-copilot-basic",
  CustomCopilotRagCustomize = "custom-copilot-rag-customize",
  CustomCopilotRagCustomApi = "custom-copilot-rag-custom-api", // TODO: handled by xxx generator
  CustomCopilotRagAzureAISearch = "custom-copilot-rag-azure-ai-search",
  CustomCopilotRagMicrosoft365 = "custom-copilot-rag-microsoft365",
  CustomCopilotAssistantNew = "custom-copilot-assistant-new",
  CustomCopilotAssistantAssistantsApi = "custom-copilot-assistant-assistants-api",

  // from TDP only
  TabAndDefaultBot = "non-sso-tab-default-bot",
  BotAndMessageExtension = "default-bot-message-extension",

  // VS only
  Empty = "empty",
  AIBot = "ai-bot",
  AIAssistantBot = "ai-assistant-bot",

  // WXP
  OutlookTaskpane = "office-addin-outlook-taskpane", // handled by OfficeAddinGeneratorNew
  WXPTaskpane = "office-addin-wxpo-taskpane", // handled by OfficeAddinGeneratorNew
  OfficeAddinCommon = "office-addin-config", // handled by OfficeAddinGeneratorNew
}

export function tryGetTemplateName(inputs: Inputs): TemplateNames | undefined {
  for (const [key, value] of inputsToTemplateName) {
    if (Object.keys(key).every((k) => key[k] === inputs[k])) {
      return value;
    }
  }
}

export function getTemplateName(inputs: Inputs): TemplateNames {
  const templateName = tryGetTemplateName(inputs);
  if (!templateName) {
    throw new Error("Template name not found");
  }
  return templateName;
}

// When multiple template name matches, only the top one will be picked.
export const inputsToTemplateName: Map<{ [key: string]: any }, TemplateNames> = new Map([
  [{ [QuestionNames.Capabilities]: CapabilityOptions.empty().id }, TemplateNames.Empty],
  [{ [QuestionNames.Capabilities]: CapabilityOptions.nonSsoTab().id }, TemplateNames.Tab],
  [
    { [QuestionNames.Capabilities]: CapabilityOptions.m365SsoLaunchPage().id },
    TemplateNames.SsoTabObo,
  ],
  [
    { [QuestionNames.Capabilities]: CapabilityOptions.dashboardTab().id },
    TemplateNames.DashboardTab,
  ],
  [
    {
      [QuestionNames.Capabilities]: CapabilityOptions.notificationBot().id,
      [QuestionNames.BotTrigger]: NotificationTriggerOptions.appService().id,
    },
    TemplateNames.NotificationExpress,
  ],
  [
    {
      [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.CSharp,
      [QuestionNames.Capabilities]: CapabilityOptions.notificationBot().id,
      [QuestionNames.BotTrigger]: NotificationTriggerOptions.appServiceForVS().id,
    },
    TemplateNames.NotificationWebApi,
  ],
  [
    {
      [QuestionNames.Capabilities]: CapabilityOptions.notificationBot().id,
      [QuestionNames.BotTrigger]: NotificationTriggerOptions.functionsHttpTrigger().id,
    },
    TemplateNames.NotificationHttpTrigger,
  ],
  [
    {
      [QuestionNames.Capabilities]: CapabilityOptions.notificationBot().id,
      [QuestionNames.BotTrigger]: NotificationTriggerOptions.functionsTimerTrigger().id,
    },
    TemplateNames.NotificationTimerTrigger,
  ],
  [
    {
      [QuestionNames.Capabilities]: CapabilityOptions.notificationBot().id,
      [QuestionNames.BotTrigger]: NotificationTriggerOptions.functionsHttpAndTimerTrigger().id,
    },
    TemplateNames.NotificationHttpTimerTrigger,
  ],
  [
    { [QuestionNames.Capabilities]: CapabilityOptions.commandBot().id },
    TemplateNames.CommandAndResponse,
  ],
  [{ [QuestionNames.Capabilities]: CapabilityOptions.workflowBot().id }, TemplateNames.Workflow],
  [{ [QuestionNames.Capabilities]: CapabilityOptions.basicBot().id }, TemplateNames.DefaultBot],
  [{ [QuestionNames.Capabilities]: CapabilityOptions.me().id }, TemplateNames.MessageExtension],
  [
    { [QuestionNames.Capabilities]: CapabilityOptions.collectFormMe().id },
    TemplateNames.MessageExtensionAction,
  ],
  [
    { [QuestionNames.Capabilities]: CapabilityOptions.SearchMe().id },
    TemplateNames.MessageExtensionSearch,
  ],
  [
    {
      [QuestionNames.Capabilities]: CapabilityOptions.m365SearchMe().id,
      [QuestionNames.MeArchitectureType]: MeArchitectureOptions.botMe().id,
    },
    TemplateNames.MessageExtensionM365,
  ],
  [
    { [QuestionNames.Capabilities]: CapabilityOptions.nonSsoTabAndBot().id },
    TemplateNames.TabAndDefaultBot,
  ],
  [
    { [QuestionNames.Capabilities]: CapabilityOptions.botAndMe().id },
    TemplateNames.BotAndMessageExtension,
  ],
  [
    { [QuestionNames.Capabilities]: CapabilityOptions.linkUnfurling().id },
    TemplateNames.LinkUnfurling,
  ],
  [{ [QuestionNames.Capabilities]: CapabilityOptions.aiBot().id }, TemplateNames.AIBot],
  [
    { [QuestionNames.Capabilities]: CapabilityOptions.aiAssistantBot().id },
    TemplateNames.AIAssistantBot,
  ],
  [
    {
      [QuestionNames.Capabilities]: CapabilityOptions.m365SearchMe().id,
      [QuestionNames.MeArchitectureType]: MeArchitectureOptions.newApi().id,
      [QuestionNames.ApiAuth]: ApiAuthOptions.none().id,
    },
    TemplateNames.CopilotPluginFromScratch,
  ],
  [
    {
      [QuestionNames.Capabilities]: CapabilityOptions.m365SearchMe().id,
      [QuestionNames.MeArchitectureType]: MeArchitectureOptions.newApi().id,
      [QuestionNames.ApiAuth]: ApiAuthOptions.bearerToken().id,
    },
    TemplateNames.CopilotPluginFromScratchApiKey,
  ],
  [
    {
      [QuestionNames.Capabilities]: CapabilityOptions.m365SearchMe().id,
      [QuestionNames.MeArchitectureType]: MeArchitectureOptions.newApi().id,
      [QuestionNames.ApiAuth]: ApiAuthOptions.microsoftEntra().id,
    },
    TemplateNames.ApiMessageExtensionSso,
  ],
  [
    { [QuestionNames.Capabilities]: CapabilityOptions.customCopilotBasic().id },
    TemplateNames.CustomCopilotBasic,
  ],
  [
    {
      [QuestionNames.Capabilities]: CapabilityOptions.customCopilotRag().id,
      [QuestionNames.CustomCopilotRag]: CustomCopilotRagOptions.customize().id,
    },
    TemplateNames.CustomCopilotRagCustomize,
  ],
  [
    {
      [QuestionNames.Capabilities]: CapabilityOptions.customCopilotRag().id,
      [QuestionNames.CustomCopilotRag]: CustomCopilotRagOptions.azureAISearch().id,
    },
    TemplateNames.CustomCopilotRagAzureAISearch,
  ],
  // [
  //   {
  //     [QuestionNames.Capabilities]: CapabilityOptions.customCopilotRag().id,
  //     [QuestionNames.CustomCopilotRag]: CustomCopilotRagOptions.customApi().id,
  //   },
  //   TemplateNames.CustomCopilotRagCustomApi,
  // ],
  [
    {
      [QuestionNames.Capabilities]: CapabilityOptions.customCopilotRag().id,
      [QuestionNames.CustomCopilotRag]: CustomCopilotRagOptions.microsoft365().id,
    },
    TemplateNames.CustomCopilotRagMicrosoft365,
  ],
  [
    {
      [QuestionNames.Capabilities]: CapabilityOptions.customCopilotAssistant().id,
      [QuestionNames.CustomCopilotAssistant]: CustomCopilotAssistantOptions.new().id,
    },
    TemplateNames.CustomCopilotAssistantNew,
  ],
  [
    {
      [QuestionNames.Capabilities]: CapabilityOptions.customCopilotAssistant().id,
      [QuestionNames.CustomCopilotAssistant]: CustomCopilotAssistantOptions.assistantsApi().id,
    },
    TemplateNames.CustomCopilotAssistantAssistantsApi,
  ],
]);

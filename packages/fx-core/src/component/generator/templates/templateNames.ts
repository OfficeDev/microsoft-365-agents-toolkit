// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { Inputs } from "@microsoft/teamsfx-api";
import {
  ApiAuthOptions,
  ApiPluginStartOptions,
  CapabilityOptions,
  CustomCopilotAssistantOptions,
  CustomCopilotRagOptions,
  DeclarativeCopilotTypeOptions,
  MeArchitectureOptions,
  NotificationTriggerOptions,
  ProgrammingLanguage,
  QuestionNames,
} from "../../../question/constants";

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
  NotificationHttpTrigger = "notification-http-trigger",
  NotificationTimerTrigger = "notification-timer-trigger",
  NotificationHttpTimerTrigger = "notification-http-timer-trigger",
  CommandAndResponse = "command-and-response",
  Workflow = "workflow",
  DefaultBot = "default-bot",
  MessageExtension = "message-extension",
  MessageExtensionAction = "message-extension-action",
  MessageExtensionSearch = "message-extension-search",
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
}

// TODO: remove this mapping after all generators are migrated to new generator pattern
export const Feature2TemplateName = {
  [`${CapabilityOptions.empty().id}:undefined`]: TemplateNames.Empty,
  [`${CapabilityOptions.nonSsoTab().id}:undefined`]: TemplateNames.Tab,
  [`${CapabilityOptions.tab().id}:undefined`]: TemplateNames.SsoTab,
  [`${CapabilityOptions.m365SsoLaunchPage().id}:undefined`]: TemplateNames.SsoTabObo,
  [`${CapabilityOptions.nonSsoTab().id}:ssr`]: TemplateNames.TabSSR,
  [`${CapabilityOptions.tab().id}:ssr`]: TemplateNames.SsoTabSSR,
  [`${CapabilityOptions.dashboardTab().id}:undefined`]: TemplateNames.DashboardTab,
  [`${CapabilityOptions.notificationBot().id}:${NotificationTriggerOptions.appService().id}`]:
    TemplateNames.NotificationExpress,
  [`${CapabilityOptions.notificationBot().id}:${NotificationTriggerOptions.appServiceForVS().id}`]:
    TemplateNames.NotificationWebApi,
  [`${CapabilityOptions.notificationBot().id}:${
    NotificationTriggerOptions.functionsHttpTrigger().id
  }`]: TemplateNames.NotificationHttpTrigger,
  [`${CapabilityOptions.notificationBot().id}:${
    NotificationTriggerOptions.functionsTimerTrigger().id
  }`]: TemplateNames.NotificationTimerTrigger,
  [`${CapabilityOptions.notificationBot().id}:${
    NotificationTriggerOptions.functionsHttpAndTimerTrigger().id
  }`]: TemplateNames.NotificationHttpTimerTrigger,
  [`${CapabilityOptions.commandBot().id}:undefined`]: TemplateNames.CommandAndResponse,
  [`${CapabilityOptions.workflowBot().id}:undefined`]: TemplateNames.Workflow,
  [`${CapabilityOptions.basicBot().id}:undefined`]: TemplateNames.DefaultBot,
  [`${CapabilityOptions.me().id}:undefined`]: TemplateNames.MessageExtension,
  [`${CapabilityOptions.collectFormMe().id}:undefined`]: TemplateNames.MessageExtensionAction,
  [`${CapabilityOptions.SearchMe().id}:undefined`]: TemplateNames.MessageExtensionSearch,
  [`${CapabilityOptions.m365SearchMe().id}:undefined:${MeArchitectureOptions.botMe().id}`]:
    TemplateNames.M365MessageExtension,
  [`${CapabilityOptions.nonSsoTabAndBot().id}:undefined`]: TemplateNames.TabAndDefaultBot,
  [`${CapabilityOptions.botAndMe().id}:undefined`]: TemplateNames.BotAndMessageExtension,
  [`${CapabilityOptions.linkUnfurling().id}:undefined`]: TemplateNames.LinkUnfurling,
  [`${CapabilityOptions.aiBot().id}:undefined`]: TemplateNames.AIBot,
  [`${CapabilityOptions.aiAssistantBot().id}:undefined`]: TemplateNames.AIAssistantBot,
  [`${CapabilityOptions.m365SearchMe().id}:undefined:${MeArchitectureOptions.newApi().id}:${
    ApiAuthOptions.none().id
  }`]: TemplateNames.CopilotPluginFromScratch,
  [`${CapabilityOptions.m365SearchMe().id}:undefined:${MeArchitectureOptions.newApi().id}:${
    ApiAuthOptions.bearerToken().id
  }`]: TemplateNames.CopilotPluginFromScratchApiKey,
  [`${CapabilityOptions.m365SearchMe().id}:undefined:${MeArchitectureOptions.newApi().id}:${
    ApiAuthOptions.microsoftEntra().id
  }`]: TemplateNames.ApiMessageExtensionSso,
  [`${CapabilityOptions.customCopilotBasic().id}:undefined`]: TemplateNames.CustomCopilotBasic,
  [`${CapabilityOptions.customCopilotRag().id}:undefined:${
    CustomCopilotRagOptions.customize().id
  }`]: TemplateNames.CustomCopilotRagCustomize,
  [`${CapabilityOptions.customCopilotRag().id}:undefined:${
    CustomCopilotRagOptions.azureAISearch().id
  }`]: TemplateNames.CustomCopilotRagAzureAISearch,
  [`${CapabilityOptions.customCopilotRag().id}:undefined:${
    CustomCopilotRagOptions.customApi().id
  }`]: TemplateNames.CustomCopilotRagCustomApi,
  [`${CapabilityOptions.customCopilotRag().id}:undefined:${
    CustomCopilotRagOptions.microsoft365().id
  }`]: TemplateNames.CustomCopilotRagMicrosoft365,
  [`${CapabilityOptions.customCopilotAssistant().id}:undefined:${
    CustomCopilotAssistantOptions.new().id
  }`]: TemplateNames.CustomCopilotAssistantNew,
  [`${CapabilityOptions.customCopilotAssistant().id}:undefined:${
    CustomCopilotAssistantOptions.assistantsApi().id
  }`]: TemplateNames.CustomCopilotAssistantAssistantsApi,
};

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
  [{ [QuestionNames.Capabilities]: CapabilityOptions.tab().id }, TemplateNames.SsoTab],
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
    TemplateNames.M365MessageExtension,
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

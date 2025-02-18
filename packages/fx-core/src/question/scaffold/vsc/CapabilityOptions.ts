// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs, OptionItem } from "@microsoft/teamsfx-api";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";
import { getLocalizedString } from "../../../common/localizeUtils";
import { TemplateNames } from "../../../component/generator/templates/templateNames";
import {
  HostType,
  HostTypeTriggerOptionItem,
  NotificationTriggers,
  QuestionNames,
} from "../../constants";

export class TabCapabilityOptions {
  static nonSsoTab(): OptionItem {
    return {
      id: "tab-non-sso",
      label: `${getLocalizedString("core.TabNonSso.label")}`,
      detail: getLocalizedString("core.TabNonSso.detail"),
      description: getLocalizedString(
        "core.createProjectQuestion.option.description.worksInOutlookM365"
      ),
      data: TemplateNames.Tab,
    };
  }
  static m365SsoLaunchPage(): OptionItem {
    return {
      id: "sso-launch-page",
      label: `${getLocalizedString("core.M365SsoLaunchPageOptionItem.label")}`,
      detail: getLocalizedString("core.M365SsoLaunchPageOptionItem.detail"),
      description: getLocalizedString(
        "core.createProjectQuestion.option.description.worksInOutlookM365"
      ),
      data: TemplateNames.SsoTabObo,
    };
  }

  static dashboardTab(): OptionItem {
    return {
      id: "dashboard-tab",
      label: getLocalizedString("core.DashboardOption.label"),
      detail: getLocalizedString("core.DashboardOption.detail"),
      description: getLocalizedString(
        "core.createProjectQuestion.option.description.worksInOutlookM365"
      ),
      data: TemplateNames.DashboardTab,
      buttons: [
        {
          iconPath: "file-symlink-file",
          tooltip: getLocalizedString("core.option.github"),
          command: "fx-extension.openTutorial",
        },
      ],
    };
  }

  // TODO: need further sub-options to decide template name
  static SPFxTab(): OptionItem {
    return {
      id: "tab-spfx",
      label: getLocalizedString("core.TabSPFxOption.labelNew"),
      description: getLocalizedString(
        "core.createProjectQuestion.option.description.worksInOutlookM365"
      ),
      detail: getLocalizedString("core.TabSPFxOption.detailNew"),
    };
  }
}

export class CustomCopilotCapabilityOptions {
  // custom copilot
  static basicChatbot(): OptionItem {
    const description = featureFlagManager.getBooleanValue(FeatureFlags.CEAEnabled)
      ? getLocalizedString("core.createProjectQuestion.capability.customEngineAgent.description")
      : undefined;
    return {
      id: "custom-copilot-basic",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotBasicOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotBasicOption.detail"
      ),
      description: description,
      data: TemplateNames.CustomCopilotBasic,
    };
  }

  static customCopilotRag(): OptionItem {
    const description = featureFlagManager.getBooleanValue(FeatureFlags.CEAEnabled)
      ? getLocalizedString("core.createProjectQuestion.capability.customEngineAgent.description")
      : undefined;
    return {
      id: "custom-copilot-rag",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagOption.detail"
      ),
      description: description,
    };
  }

  static aiAgent(): OptionItem {
    const description = featureFlagManager.getBooleanValue(FeatureFlags.CEAEnabled)
      ? getLocalizedString("core.createProjectQuestion.capability.customEngineAgent.description")
      : undefined;
    return {
      id: "custom-copilot-agent",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotAssistantOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotAssistantOption.detail"
      ),
      description: description,
    };
  }
}

export class CustomCopilotRagOptions {
  static customize(): OptionItem {
    return {
      id: "custom-copilot-rag-customize",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagCustomizeOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagCustomizeOption.detail"
      ),
      data: TemplateNames.CustomCopilotRagCustomize,
    };
  }

  static azureAISearch(): OptionItem {
    return {
      id: "custom-copilot-rag-azureAISearch",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagAzureAISearchOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagAzureAISearchOption.detail"
      ),
      data: TemplateNames.CustomCopilotRagAzureAISearch,
    };
  }

  static customApi(): OptionItem {
    return {
      id: "custom-copilot-rag-customApi",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagCustomApiOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagCustomApiOption.detail"
      ),
      description: getLocalizedString("core.createProjectQuestion.option.description.preview"),
      data: TemplateNames.CustomCopilotRagCustomApi,
    };
  }

  static microsoft365(): OptionItem {
    return {
      id: "custom-copilot-rag-microsoft365",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagMicrosoft365Option.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotRagMicrosoft365Option.detail"
      ),
      data: TemplateNames.CustomCopilotRagMicrosoft365,
    };
  }
}

export class CustomCopilotAssistantOptions {
  static new(): OptionItem {
    return {
      id: "custom-copilot-agent-new",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotAssistantNewOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotAssistantNewOption.detail"
      ),
      data: TemplateNames.CustomCopilotAssistantNew,
    };
  }

  static assistantsApi(): OptionItem {
    return {
      id: "custom-copilot-agent-assistants-api",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotAssistantAssistantsApiOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotAssistantAssistantsApiOption.detail"
      ),
      description: getLocalizedString("core.createProjectQuestion.option.description.preview"),
      data: TemplateNames.CustomCopilotAssistantAssistantsApi,
    };
  }
}

export class MeCapabilityOptions {
  // need further sub-options to decide template name
  static m365SearchMe(): OptionItem {
    return {
      id: "search-app",
      label: `${getLocalizedString("core.M365SearchAppOptionItem.label")}`,
      detail: getLocalizedString("core.M365SearchAppOptionItem.copilot.detail"),
    };
  }

  static collectFormMe(): OptionItem {
    return {
      id: "collect-form-message-extension",
      label: `${getLocalizedString("core.MessageExtensionOption.labelNew")}`,
      detail: getLocalizedString("core.MessageExtensionOption.detail"),
      data: TemplateNames.MessageExtensionAction,
    };
  }

  static linkUnfurling(): OptionItem {
    return {
      id: "link-unfurling",
      label: `${getLocalizedString("core.LinkUnfurlingOption.label")}`,
      detail: getLocalizedString("core.LinkUnfurlingOption.detail"),
      description: getLocalizedString(
        "core.createProjectQuestion.option.description.worksInOutlook"
      ),
      data: TemplateNames.LinkUnfurling,
    };
  }
}

export class MeArchitectureOptions {
  static botMe(): OptionItem {
    return {
      id: "bot",
      label: getLocalizedString("core.createProjectQuestion.capability.botMessageExtension.label"),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.botMessageExtension.detail"
      ),
      description: getLocalizedString(
        "core.createProjectQuestion.option.description.worksInOutlook"
      ),
      data: TemplateNames.MessageExtensionM365,
    };
  }

  static newApi(): OptionItem {
    return {
      id: "new-api",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.copilotPluginNewApiOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.messageExtensionNewApiOption.detail"
      ),
    };
  }

  static openApiSpec(): OptionItem {
    return {
      id: "api-spec",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.copilotPluginApiSpecOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.messageExtensionApiSpecOption.detail"
      ),
      data: TemplateNames.MessageExtensionWithExistingApiSpec,
    };
  }
}

export class OfficeAddinCapabilityOptions {
  static outlookAddinImport(): OptionItem {
    return {
      id: "outlook-addin-import",
      label: getLocalizedString("core.importAddin.label"),
      detail: getLocalizedString("core.importAddin.detail"),
      data: TemplateNames.OfficeAddinCommon,
    };
  }
  static officeContentAddin(): OptionItem {
    return {
      id: "office-content-addin",
      label: getLocalizedString("core.officeContentAddin.label"),
      detail: getLocalizedString("core.officeContentAddin.detail"),
    };
  }
  static officeAddinImport(): OptionItem {
    return {
      id: "office-addin-import",
      label: getLocalizedString("core.importOfficeAddin.label"),
      detail: getLocalizedString("core.importAddin.detail"),
      description: getLocalizedString(
        "core.createProjectQuestion.option.description.previewOnWindow"
      ),
      data: TemplateNames.OfficeAddinCommon,
    };
  }
  static outlookTaskPane(): OptionItem {
    return {
      id: "outlook-json-taskpane",
      label: getLocalizedString("core.newTaskpaneAddin.label"),
      detail: getLocalizedString("core.newTaskpaneAddin.detail"),
      data: TemplateNames.OutlookTaskpane,
    };
  }
  static wxpTaskPane(): OptionItem {
    return {
      id: "wxp-json-taskpane",
      label: getLocalizedString("core.newTaskpaneAddin.label"),
      detail: getLocalizedString("core.newTaskpaneAddin.detail"),
      data: TemplateNames.WXPTaskpane,
    };
  }
}

export class BotCapabilityOptions {
  static readonly basicBotId = "bot";
  static readonly notificationBotId = "notification";
  static readonly commandBotId = "command-bot";
  static readonly workflowBotId = "workflow-bot";

  static basicBot(): OptionItem {
    return {
      id: BotCapabilityOptions.basicBotId,
      label: `${getLocalizedString("core.BotNewUIOption.label")}`,
      detail: getLocalizedString("core.BotNewUIOption.detail"),
      data: TemplateNames.DefaultBot,
    };
  }
  static aiBot(): OptionItem {
    return {
      id: "ai-bot",
      label: getLocalizedString("core.aiBotOption.label"),
      detail: getLocalizedString("core.aiBotOption.detail"),
      data: TemplateNames.AIBot,
    };
  }
  // need further sub-options to decide template name
  static notificationBot(): OptionItem {
    return {
      id: BotCapabilityOptions.notificationBotId,
      label: `${getLocalizedString("core.NotificationOption.label")}`,
      detail: getLocalizedString("core.NotificationOption.detail"),
      buttons: [
        {
          iconPath: "file-symlink-file",
          tooltip: getLocalizedString("core.option.github"),
          command: "fx-extension.openTutorial",
        },
      ],
    };
  }
  static commandBot(): OptionItem {
    return {
      // id must match cli `yargsHelp`
      id: BotCapabilityOptions.commandBotId,
      label: `${getLocalizedString("core.CommandAndResponseOption.label")}`,
      detail: getLocalizedString("core.CommandAndResponseOption.detail"),
      data: TemplateNames.CommandAndResponse,
      buttons: [
        {
          iconPath: "file-symlink-file",
          tooltip: getLocalizedString("core.option.github"),
          command: "fx-extension.openTutorial",
        },
      ],
    };
  }

  static workflowBot(): OptionItem {
    const item: OptionItem = {
      id: BotCapabilityOptions.workflowBotId,
      label: `${getLocalizedString("core.WorkflowOption.label")}`,
      detail: getLocalizedString("core.WorkflowOption.detail"),
      data: TemplateNames.Workflow,
      buttons: [
        {
          iconPath: "file-symlink-file",
          tooltip: getLocalizedString("core.option.github"),
          command: "fx-extension.openTutorial",
        },
      ],
    };
    return item;
  }
}

export class NotificationBotOptions {
  static appService(): HostTypeTriggerOptionItem {
    return {
      id: "http-express",
      hostType: HostType.AppService,
      label: getLocalizedString("plugins.bot.triggers.http-express.label"),
      description: getLocalizedString("plugins.bot.triggers.http-express.description"),
      detail: getLocalizedString("plugins.bot.triggers.http-express.detail"),
      data: TemplateNames.NotificationExpress,
    };
  }
  static appServiceForVS(): HostTypeTriggerOptionItem {
    return {
      id: "http-webapi",
      hostType: HostType.AppService,
      label: getLocalizedString("plugins.bot.triggers.http-webapi.label"),
      description: getLocalizedString("plugins.bot.triggers.http-webapi.description"),
      detail: getLocalizedString("plugins.bot.triggers.http-webapi.detail"),
      data: TemplateNames.NotificationWebApi,
    };
  }
  static functionsTimerTrigger(): HostTypeTriggerOptionItem {
    return {
      id: "timer-functions",
      hostType: HostType.Functions,
      triggers: [NotificationTriggers.TIMER],
      label: getLocalizedString("plugins.bot.triggers.timer-functions.label"),
      description: getLocalizedString("plugins.bot.triggers.timer-functions.description"),
      detail: getLocalizedString("plugins.bot.triggers.timer-functions.detail"),
      data: TemplateNames.NotificationTimerTrigger,
    };
  }

  static functionsHttpAndTimerTrigger(): HostTypeTriggerOptionItem {
    return {
      id: "http-and-timer-functions",
      hostType: HostType.Functions,
      triggers: [NotificationTriggers.HTTP, NotificationTriggers.TIMER],
      label: getLocalizedString("plugins.bot.triggers.http-and-timer-functions.label"),
      description: getLocalizedString("plugins.bot.triggers.http-and-timer-functions.description"),
      detail: getLocalizedString("plugins.bot.triggers.http-and-timer-functions.detail"),
      data: TemplateNames.NotificationHttpTimerTrigger,
    };
  }

  static functionsHttpTrigger(): HostTypeTriggerOptionItem {
    return {
      id: "http-functions",
      hostType: HostType.Functions,
      triggers: [NotificationTriggers.HTTP],
      label: getLocalizedString("plugins.bot.triggers.http-functions.label"),
      description: getLocalizedString("plugins.bot.triggers.http-functions.description"),
      detail: getLocalizedString("plugins.bot.triggers.http-functions.detail"),
      data: TemplateNames.NotificationHttpTrigger,
    };
  }
}

export class DACapabilityOptions {
  static declarativeAgent(): OptionItem {
    return {
      id: "declarative-agent",
      label: getLocalizedString("core.createProjectQuestion.projectType.declarativeAgent.label"),
      detail: getLocalizedString("core.createProjectQuestion.projectType.declarativeAgent.detail"),
    };
  }
  static noPlugin(): OptionItem {
    return {
      id: "no",
      label: getLocalizedString("core.createProjectQuestion.noPlugin.label"),
      detail: getLocalizedString("core.createProjectQuestion.noPlugin.detail"),
      data: TemplateNames.BasicGpt,
    };
  }
  static withPlugin(): OptionItem {
    return {
      id: "yes",
      label: getLocalizedString("core.createProjectQuestion.addPlugin.label"),
      detail: getLocalizedString("core.createProjectQuestion.addPlugin.detail"),
    };
  }
}

export class ApiPluginStartOptions {
  static newApi(): OptionItem {
    return {
      id: "new-api",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.copilotPluginNewApiOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.copilotPluginNewApiOption.detail"
      ),
    };
  }

  static apiSpec(): OptionItem {
    return {
      id: "api-spec",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.copilotPluginApiSpecOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.copilotPluginApiSpecOption.detail"
      ),
      data: TemplateNames.ApiPluginWithExistingApiSpec,
    };
  }

  static existingPlugin(): OptionItem {
    return {
      id: "existing-plugin",
      label: getLocalizedString("core.createProjectQuestion.apiPlugin.importPlugin.label"),
      detail: getLocalizedString("core.createProjectQuestion.apiPlugin.importPlugin.detail"),
      data: TemplateNames.BasicGpt,
    };
  }
}

export class ApiAuthOptions {
  static none(isME = false): OptionItem {
    return {
      id: "none",
      label: "None",
      data: isME
        ? TemplateNames.MessageExtensionWithNewApiFromScratch
        : TemplateNames.ApiPluginFromScratch,
    };
  }
  static apiKey(): OptionItem {
    return {
      id: "api-key",
      label: "API Key",
      // TODO: Update the name because the ApiPluginFromScratchBearer template is currently actually ApiPluginFromScratchAPIKey.
      data: TemplateNames.ApiPluginFromScratchBearer,
    };
  }
  static bearerToken(): OptionItem {
    return {
      id: "bearer-token",
      label: "API Key (Bearer Token Auth)",
      data: TemplateNames.MessageExtensionWithNewApiFromScratchUsingApiKey,
    };
  }
  static microsoftEntra(isME = false): OptionItem {
    return {
      id: "microsoft-entra",
      label: "Microsoft Entra",
      data: isME
        ? TemplateNames.MessageExtensionWithNewApiFromScratchUsingOAuth
        : TemplateNames.ApiPluginFromScratchOAuth,
    };
  }

  static oauth(): OptionItem {
    return {
      id: "oauth",
      label: "OAuth",
      data: TemplateNames.ApiPluginFromScratchOAuth,
    };
  }
}

export function setTemplateName(selected: string | OptionItem, inputs: Inputs): void {
  if ((selected as OptionItem).data) {
    inputs[QuestionNames.TemplateName] = (selected as OptionItem).data as string;
  }
}

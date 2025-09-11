// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { IQTreeNode, OptionItem, Platform } from "@microsoft/teamsfx-api";
import { getLocalizedString } from "../../../common/localizeUtils";
import { TemplateNames } from "../../../component/generator/templates/templateNames";
import { appNameQuestion, folderQuestion } from "../../create";
import { QuestionNames } from "../../questionNames";
import {
  BotCapabilityOptions,
  MeCapabilityOptions,
  setTemplateName,
  TeamsAgentCapabilityOptions,
} from "../vsc/CapabilityOptions";
import { folderAndAppNameCondition, languageNode } from "../vsc/createRootNode";
import { daProjectTypeNode } from "../vsc/daProjectTypeNode";
import {
  customCopilotRagNode,
  llmServiceNode,
  m365SearchMeSubNode,
  notificationBotTriggerNode,
} from "../vsc/teamsProjectTypeNode";

export class VSCapabilityOptions {
  // empty
  static empty(): OptionItem {
    return {
      id: "empty",
      label: "Empty",
      data: TemplateNames.Empty,
    };
  }
  static declarativeAgent(): OptionItem {
    return {
      id: "declarative-agent",
      label: getLocalizedString("core.createProjectQuestion.projectType.declarativeAgent.label"),
      detail: getLocalizedString("core.createProjectQuestion.projectType.declarativeAgent.detail"),
    };
  }
  static nonSsoTab(): OptionItem {
    return {
      id: "tab-non-sso",
      label: `${getLocalizedString("core.TabNonSso.label")}`,
      detail: getLocalizedString("core.TabNonSso.detail"),
      description: getLocalizedString(
        "core.createProjectQuestion.option.description.worksInOutlookM365"
      ),
      data: TemplateNames.TabSSR,
    };
  }
  static tab(): OptionItem {
    return {
      id: "tab",
      label: getLocalizedString("core.TabOption.label"),
      description: getLocalizedString("core.TabOption.description"),
      detail: getLocalizedString("core.TabOption.detail"),
      data: TemplateNames.SsoTabSSR,
    };
  }
  static SearchMeVS(): OptionItem {
    return {
      id: "search-message-extension",
      label: `${getLocalizedString("core.M365SearchAppOptionItem.label")}`,
      detail: getLocalizedString("core.SearchAppOptionItem.detail"),
      data: TemplateNames.MessageExtensionSearch,
    };
  }
  // Currently only for dotnet
  static weatherAgentBot(): OptionItem {
    return {
      id: "custom-copilot-weather-agent",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotWeatherOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotWeatherOption.detail"
      ),
      data: TemplateNames.WeatherAgent,
    };
  }
}

/**
 * Scaffold question model dedicated for VS platform
 */

export function scaffoldQuestionForVS(): IQTreeNode {
  const node: IQTreeNode = {
    data: { type: "group" },
    children: [
      {
        data: {
          name: QuestionNames.Capabilities,
          title: getLocalizedString("core.createCapabilityQuestion.titleNew"),
          type: "singleSelect",
          staticOptions: [
            VSCapabilityOptions.empty(),
            VSCapabilityOptions.declarativeAgent(),
            TeamsAgentCapabilityOptions.basicChatbot(),
            TeamsAgentCapabilityOptions.customCopilotRag(),
            // TeamsAgentCapabilityOptions.aiAgent(),
            VSCapabilityOptions.weatherAgentBot(),
            BotCapabilityOptions.basicBot(),
            BotCapabilityOptions.notificationBot(),
            BotCapabilityOptions.commandBot(),
            BotCapabilityOptions.workflowBot(),
            VSCapabilityOptions.nonSsoTab(),
            VSCapabilityOptions.tab(),
            MeCapabilityOptions.m365SearchMe(),
            MeCapabilityOptions.collectFormMe(),
            VSCapabilityOptions.SearchMeVS(),
            MeCapabilityOptions.linkUnfurling(),
          ],
          onDidSelection: setTemplateName,
        },
        children: [
          daProjectTypeNode(VSCapabilityOptions.declarativeAgent().id),
          customCopilotRagNode(),
          // aiAgentNode(),
          m365SearchMeSubNode(),
          llmServiceNode({
            enum: [
              TeamsAgentCapabilityOptions.basicChatbot().id,
              TeamsAgentCapabilityOptions.customCopilotRag().id,
              // TeamsAgentCapabilityOptions.aiAgent().id,
              VSCapabilityOptions.weatherAgentBot().id,
            ],
          }),
          notificationBotTriggerNode(Platform.VS),
        ],
      },
      languageNode(),
      {
        condition: folderAndAppNameCondition,
        data: {
          type: "group",
        },
        children: [
          {
            data: folderQuestion(),
          },
          {
            data: appNameQuestion(),
          },
        ],
      },
    ],
  };
  return node;
}

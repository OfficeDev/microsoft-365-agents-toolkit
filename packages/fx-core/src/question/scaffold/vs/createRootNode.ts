// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { IQTreeNode, OptionItem } from "@microsoft/teamsfx-api";
import { getLocalizedString } from "../../../common/localizeUtils";
import { TemplateNames } from "../../../component/generator/templates/templateNames";
import { appNameQuestion, folderQuestion } from "../../create";
import { QuestionNames } from "../../questionNames";
import { llmServiceNode } from "../commonNodes";
import {
  BotCapabilityOptions,
  MeCapabilityOptions,
  setTemplateName,
  TeamsAgentCapabilityOptions,
} from "../vsc/CapabilityOptions";
import { folderAndAppNameCondition, languageNode } from "../vsc/createRootNode";
import { daProjectTypeNode } from "../vsc/daProjectTypeNode";
import { customCopilotRagNode } from "../vsc/teamsProjectTypeNode";

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
      data: TemplateNames.Tab,
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

  static travelAgentBot(): OptionItem {
    return {
      id: "custom-copilot-travel-agent",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotTravelOption.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.customCopilotTravelOption.detail"
      ),
      data: TemplateNames.TravelAgent,
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
            TeamsAgentCapabilityOptions.collaboratorAgent(),
            TeamsAgentCapabilityOptions.customCopilotRag(),
            VSCapabilityOptions.weatherAgentBot(),
            VSCapabilityOptions.travelAgentBot(),
            BotCapabilityOptions.basicBot(),
            VSCapabilityOptions.nonSsoTab(),
            MeCapabilityOptions.basicMe(),
          ],
          onDidSelection: setTemplateName,
        },
        children: [
          daProjectTypeNode(VSCapabilityOptions.declarativeAgent().id),
          customCopilotRagNode(),
          llmServiceNode({
            enum: [
              TeamsAgentCapabilityOptions.basicChatbot().id,
              TeamsAgentCapabilityOptions.collaboratorAgent().id,
              TeamsAgentCapabilityOptions.customCopilotRag().id,
              VSCapabilityOptions.weatherAgentBot().id,
              VSCapabilityOptions.travelAgentBot().id,
            ],
          }),
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

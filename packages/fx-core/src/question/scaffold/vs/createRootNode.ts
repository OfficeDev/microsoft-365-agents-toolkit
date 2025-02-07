// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { IQTreeNode, OptionItem, Platform } from "@microsoft/teamsfx-api";
import { getLocalizedString } from "../../../common/localizeUtils";
import { TemplateNames } from "../../../component/generator/templates/templateNames";
import { QuestionNames } from "../../constants";
import { appNameQuestion, folderQuestion } from "../../create";
import {
  BotCapabilityOptions,
  CustomCopilotCapabilityOptions,
  MeCapabilityOptions,
  setTemplateName,
} from "../vsc/CapabilityOptions";
import { folderAndAppNameCondition, languageNode } from "../vsc/createRootNode";
import { llmServiceNode } from "../vsc/customAgentProjectTypeNode";
import { daProjectTypeNode } from "../vsc/daProjectTypeNode";
import { notificationBotTriggerNode } from "../vsc/teamsProjectTypeNode";

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
  static aiAssistantBot(): OptionItem {
    return {
      id: "ai-assistant-bot",
      label: getLocalizedString("core.aiAssistantBotOption.label"),
      detail: getLocalizedString("core.aiAssistantBotOption.detail"),
      description: getLocalizedString("core.createProjectQuestion.option.description.preview"),
      data: TemplateNames.AIAssistantBot,
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
            CustomCopilotCapabilityOptions.basicChatbot(),
            CustomCopilotCapabilityOptions.customCopilotRag(),
            CustomCopilotCapabilityOptions.aiAgent(),
            BotCapabilityOptions.basicBot(),
            BotCapabilityOptions.aiBot(),
            VSCapabilityOptions.aiAssistantBot(),
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
          llmServiceNode({
            enum: [
              CustomCopilotCapabilityOptions.basicChatbot().id,
              CustomCopilotCapabilityOptions.customCopilotRag().id,
              CustomCopilotCapabilityOptions.aiAgent().id,
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

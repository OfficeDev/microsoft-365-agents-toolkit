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
import { botTriggerNode } from "../vsc/m365ProjectTypeNode";

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
}

/**
 *
 * FxCore API for scaffold: scaffold(questionModel: IQTreeNode, generators: DefaultTemplateGenerator[]): Promise<Result<any, FxError>>
 * Dedicated for VS Code platform
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
            CustomCopilotCapabilityOptions.customCopilotBasic(),
            CustomCopilotCapabilityOptions.customCopilotRag(),
            CustomCopilotCapabilityOptions.customCopilotAssistant(),
            BotCapabilityOptions.basicBot(),
            BotCapabilityOptions.aiBot(),
            BotCapabilityOptions.aiAssistantBot(),
            BotCapabilityOptions.notificationBot(),
            BotCapabilityOptions.commandBot(),
            BotCapabilityOptions.workflowBot(),
            VSCapabilityOptions.nonSsoTab(),
            VSCapabilityOptions.tab(),
            MeCapabilityOptions.m365SearchMe(),
            MeCapabilityOptions.collectFormMe(),
            MeCapabilityOptions.SearchMeVS(),
            MeCapabilityOptions.linkUnfurling(),
          ],
          onDidSelection: setTemplateName,
        },
        children: [
          daProjectTypeNode(VSCapabilityOptions.declarativeAgent().id),
          llmServiceNode({
            enum: [
              CustomCopilotCapabilityOptions.customCopilotBasic().id,
              CustomCopilotCapabilityOptions.customCopilotRag().id,
              CustomCopilotCapabilityOptions.customCopilotAssistant().id,
            ],
          }),
          botTriggerNode(Platform.VS),
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

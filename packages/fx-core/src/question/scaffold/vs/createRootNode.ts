// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { IQTreeNode, Platform } from "@microsoft/teamsfx-api";
import { getLocalizedString } from "../../../common/localizeUtils";
import { QuestionNames } from "../../constants";
import { appNameQuestion, folderQuestion } from "../../create";
import {
  BotCapabilityOptions,
  CustomCopilotCapabilityOptions,
  MeCapabilityOptions,
  setTemplateName,
  TdpCapabilityOptions,
} from "../vsc/CapabilityOptions";
import { folderAndAppNameCondition, languageNode } from "../vsc/createRootNode";
import {
  aiAgentNode,
  customCopilotRagNode,
  llmServiceNode,
} from "../vsc/customAgentProjectTypeNode";
import { daProjectTypeNode } from "../vsc/daProjectTypeNode";
import { m365SearchMeSubNode, notificationBotTriggerNode } from "../vsc/teamsProjectTypeNode";
import { VSCapabilityOptions } from "./CapabilityOptions";

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
            TdpCapabilityOptions.me(),
          ],
          onDidSelection: setTemplateName,
        },
        children: [
          daProjectTypeNode(VSCapabilityOptions.declarativeAgent().id),
          customCopilotRagNode(),
          aiAgentNode(),
          m365SearchMeSubNode(),
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

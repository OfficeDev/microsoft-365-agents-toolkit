// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs, IQTreeNode, OptionItem, Platform } from "@microsoft/teamsfx-api";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";
import { getLocalizedString } from "../../../common/localizeUtils";
import { QuestionNames } from "../../constants";
import { TemplateNames } from "../../templates";
import {
  ApiPluginStartOptions,
  BotCapabilityOptions,
  CustomCopilotCapabilityOptions,
  MeCapabilityOptions,
  setTemplateName,
  TabCapabilityOptions,
} from "../vsc/CapabilityOptions";
import { appNameNode, folderNode, languageNode } from "../vsc/createRootNode";
import { llmServiceNode } from "../vsc/customAgentProjectTypeNode";
import { declarativeAgentProjectTypeNode } from "../vsc/daProjectTypeNode";
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
            TabCapabilityOptions.nonSsoTab(),
            TabCapabilityOptions.tab(),
            MeCapabilityOptions.m365SearchMe(),
            MeCapabilityOptions.collectFormMe(),
            MeCapabilityOptions.SearchMeVS(),
            MeCapabilityOptions.linkUnfurling(),
          ],
          onDidSelection: setTemplateName,
        },
        children: [
          declarativeAgentProjectTypeNode(VSCapabilityOptions.declarativeAgent().id),
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
        condition: (inputs: Inputs) => {
          // Only skip this project when need to rediect to Kiota: 1. Feature flag enabled 2. Creating plugin/declarative copilot from existing spec 3. No plugin manifest path
          return !(
            featureFlagManager.getBooleanValue(FeatureFlags.KiotaIntegration) &&
            inputs[QuestionNames.ApiPluginType] === ApiPluginStartOptions.apiSpec().id &&
            inputs[QuestionNames.Capabilities] === VSCapabilityOptions.declarativeAgent().id &&
            !inputs[QuestionNames.ApiPluginManifestPath]
          );
        },
        data: {
          type: "group",
        },
        children: [folderNode(), appNameNode()],
      },
    ],
  };
  return node;
}

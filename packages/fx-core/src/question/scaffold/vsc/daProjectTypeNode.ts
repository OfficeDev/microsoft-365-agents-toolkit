// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { IQTreeNode, OptionItem } from "@microsoft/teamsfx-api";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";
import { getLocalizedString } from "../../../common/localizeUtils";
import { ProgrammingLanguage } from "../../constants";
import { QuestionNames } from "../../questionNames";
import { apiSpecWithSearchNode } from "../commonNodes";
import {
  ActionStartOptions,
  ApiAuthOptions,
  DACapabilityOptions,
  setTemplateName,
  setTemplateNameAndGC,
} from "./CapabilityOptions";
import { ProjectTypeOptions } from "./ProjectTypeOptions";
import { MCPServerTypeNode } from "./teamsProjectTypeNode";

export function daProjectTypeNode(
  parentValue = ProjectTypeOptions.copilotAgentOptionId
): IQTreeNode {
  return {
    // project-type = Declarative Agent
    condition: { equals: parentValue },
    data: {
      name: QuestionNames.Capabilities,
      title: getLocalizedString(
        "template.createProjectQuestion.projectType.copilotExtension.title"
      ),
      placeholder: getLocalizedString(
        "template.createProjectQuestion.projectType.copilotExtension.placeholder"
      ),
      type: "singleSelect",
      staticOptions: [DACapabilityOptions.declarativeAgent()],
      skipSingleOption: true,
    },
    children: [
      {
        condition: { equals: DACapabilityOptions.declarativeAgent().id },
        data: {
          name: QuestionNames.WithPlugin,
          title: getLocalizedString("template.createProjectQuestion.declarativeCopilot.title"),
          cliDescription: "Whether to add API plugin for your declarative Copilot.",
          type: "singleSelect",
          staticOptions: DACapabilityOptions.all(),
          placeholder: getLocalizedString(
            "template.createProjectQuestion.declarativeCopilot.placeholder"
          ),
          onDidSelection: setTemplateNameAndGC,
        },
        children: [
          {
            condition: { equals: DACapabilityOptions.withPlugin().id },
            data: {
              type: "singleSelect",
              name: QuestionNames.ActionType,
              title: getLocalizedString("template.createProjectQuestion.createApiPlugin.title"),
              cliDescription: "API plugin type.",
              placeholder: getLocalizedString(
                "template.createProjectQuestion.addApiPlugin.placeholder"
              ),
              staticOptions: [
                ActionStartOptions.newApi(),
                ActionStartOptions.apiSpecWithSearch(),
                ...(featureFlagManager.getBooleanValue(FeatureFlags.DAMetaOS)
                  ? [ActionStartOptions.DAMetaOS()]
                  : []),
                ...(featureFlagManager.getBooleanValue(FeatureFlags.MCPForDA)
                  ? [ActionStartOptions.mcp()]
                  : []),
              ],
              default: ActionStartOptions.newApi().id,
              onDidSelection: setTemplateName,
            },
            children: [
              {
                condition: { equals: ActionStartOptions.newApi().id },
                data: {
                  type: "singleSelect",
                  name: QuestionNames.ApiAuth,
                  title: getLocalizedString(
                    "template.createProjectQuestion.apiMessageExtensionAuth.title"
                  ),
                  cliDescription: "The authentication type for the API.",
                  placeholder: getLocalizedString(
                    "template.createProjectQuestion.apiMessageExtensionAuth.placeholder"
                  ),
                  staticOptions: [
                    ApiAuthOptions.none(false),
                    ApiAuthOptions.apiKey(),
                    ApiAuthOptions.microsoftEntra(),
                    ApiAuthOptions.oauth(),
                  ],
                  default: ApiAuthOptions.none().id,
                  onDidSelection: setTemplateName,
                },
              },
              apiSpecWithSearchNode(),
              MCPServerTypeNode(),
            ],
          },
        ],
      },
    ],
  };
}

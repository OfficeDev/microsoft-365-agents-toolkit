// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs, IQTreeNode } from "@microsoft/teamsfx-api";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";
import { getLocalizedString } from "../../../common/localizeUtils";
import { QuestionNames } from "../../constants";
import { pluginApiSpecQuestion, pluginManifestQuestion } from "../../create";
import {
  ApiAuthOptions,
  ApiPluginStartOptions,
  DACapabilityOptions,
  setTemplateName,
} from "./CapabilityOptions";
import { ProjectTypeOptions } from "./ProjectTypeOptions";
import { apiSpecNode } from "./teamsProjectTypeNode";

export function daProjectTypeNode(
  parentValue = ProjectTypeOptions.copilotAgentOptionId
): IQTreeNode {
  return {
    // project-type = Declarative Agent
    condition: { equals: parentValue },
    data: {
      name: QuestionNames.Capabilities,
      title: getLocalizedString("core.createProjectQuestion.projectType.copilotExtension.title"),
      placeholder: getLocalizedString(
        "core.createProjectQuestion.projectType.copilotExtension.placeholder"
      ),
      type: "singleSelect",
      staticOptions: [DACapabilityOptions.declarativeAgent()],
      skipSingleOption: true,
    },
    children: [
      {
        data: {
          name: QuestionNames.WithPlugin,
          title: getLocalizedString("core.createProjectQuestion.declarativeCopilot.title"),
          cliDescription: "Whether to add API plugin for your declarative Copilot.",
          type: "singleSelect",
          staticOptions: [DACapabilityOptions.noPlugin(), DACapabilityOptions.withPlugin()],
          placeholder: getLocalizedString(
            "core.createProjectQuestion.declarativeCopilot.placeholder"
          ),
          onDidSelection: setTemplateName,
        },
        children: [
          {
            condition: { equals: DACapabilityOptions.withPlugin().id },
            data: {
              type: "singleSelect",
              name: QuestionNames.ApiPluginType,
              title: getLocalizedString("core.createProjectQuestion.createApiPlugin.title"),
              cliDescription: "API plugin type.",
              placeholder: getLocalizedString(
                "core.createProjectQuestion.addApiPlugin.placeholder"
              ),
              staticOptions: [
                ApiPluginStartOptions.newApi(),
                ApiPluginStartOptions.apiSpec(),
                ApiPluginStartOptions.existingPlugin(),
              ],
              default: ApiPluginStartOptions.newApi().id,
              onDidSelection: setTemplateName,
            },
            children: [
              {
                condition: { equals: ApiPluginStartOptions.newApi().id },
                data: {
                  type: "singleSelect",
                  name: QuestionNames.ApiAuth,
                  title: getLocalizedString(
                    "core.createProjectQuestion.apiMessageExtensionAuth.title"
                  ),
                  cliDescription: "The authentication type for the API.",
                  placeholder: getLocalizedString(
                    "core.createProjectQuestion.apiMessageExtensionAuth.placeholder"
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
              apiSpecNode(
                (inputs: Inputs) =>
                  inputs[QuestionNames.ApiPluginType] === ApiPluginStartOptions.apiSpec().id &&
                  !featureFlagManager.getBooleanValue(FeatureFlags.KiotaIntegration)
              ),
              {
                condition: { equals: ApiPluginStartOptions.existingPlugin().id },
                data: { type: "group", name: QuestionNames.ImportPlugin },
                children: [
                  {
                    data: pluginManifestQuestion(),
                  },
                  {
                    data: pluginApiSpecQuestion(),
                  },
                ],
              },
            ],
          },
        ],
      },
    ],
  };
}

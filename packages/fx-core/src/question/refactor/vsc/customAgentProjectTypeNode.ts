// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs, IQTreeNode } from "@microsoft/teamsfx-api";
import { getLocalizedString } from "../../../common/localizeUtils";
import { QuestionNames } from "../../constants";
import {
  CustomCopilotAssistantOptions,
  CustomCopilotCapabilityOptions,
  CustomCopilotRagOptions,
  setTemplateName,
} from "./CapabilityOptions";
import { ProjectTypeOptions } from "./ProjectTypeOptions";
import { apiSpecNode } from "./m365ProjectTypeNode";

export function customEngineAgentProjectTypeNode(): IQTreeNode {
  return {
    // project-type = Custom Engine Agent
    condition: { equals: ProjectTypeOptions.customCopilotOptionId },
    data: {
      name: QuestionNames.Capabilities,
      title: getLocalizedString("core.createProjectQuestion.projectType.customCopilot.title"),
      type: "singleSelect",
      staticOptions: [
        CustomCopilotCapabilityOptions.customCopilotBasic(),
        CustomCopilotCapabilityOptions.customCopilotRag(),
        CustomCopilotCapabilityOptions.customCopilotAssistant(),
      ],
      placeholder: getLocalizedString(
        "core.createProjectQuestion.projectType.customCopilot.placeholder"
      ),
      onDidSelection: setTemplateName,
    },
    children: [
      {
        condition: { equals: CustomCopilotCapabilityOptions.customCopilotRag().id },
        data: {
          type: "singleSelect",
          name: QuestionNames.CustomCopilotRag,
          title: getLocalizedString("core.createProjectQuestion.capability.customCopilotRag.title"),
          placeholder: getLocalizedString(
            "core.createProjectQuestion.capability.customCopilotRag.placeholder"
          ),
          staticOptions: [
            CustomCopilotRagOptions.customize(),
            CustomCopilotRagOptions.azureAISearch(),
            CustomCopilotRagOptions.customApi(),
            CustomCopilotRagOptions.microsoft365(),
          ],
          default: CustomCopilotRagOptions.customize().id,
        },
        children: [apiSpecNode({ equals: CustomCopilotRagOptions.customApi().id })],
      },
      {
        condition: { equals: CustomCopilotCapabilityOptions.customCopilotAssistant().id },
        data: {
          type: "singleSelect",
          name: QuestionNames.CustomCopilotAssistant,
          title: getLocalizedString(
            "core.createProjectQuestion.capability.customCopilotAssistant.title"
          ),
          placeholder: getLocalizedString(
            "core.createProjectQuestion.capability.customCopilotAssistant.placeholder"
          ),
          staticOptions: [
            CustomCopilotAssistantOptions.new(),
            CustomCopilotAssistantOptions.assistantsApi(),
          ],
          default: CustomCopilotAssistantOptions.new().id,
        },
      },
      {
        data: {
          type: "singleSelect",
          name: QuestionNames.LLMService,
          title: getLocalizedString("core.createProjectQuestion.llmService.title"),
          placeholder: getLocalizedString("core.createProjectQuestion.llmService.placeholder"),
          staticOptions: [
            {
              id: "llm-service-azure-openai",
              label: getLocalizedString(
                "core.createProjectQuestion.llmServiceAzureOpenAIOption.label"
              ),
              detail: getLocalizedString(
                "core.createProjectQuestion.llmServiceAzureOpenAIOption.detail"
              ),
            },
            {
              id: "llm-service-openai",
              label: getLocalizedString("core.createProjectQuestion.llmServiceOpenAIOption.label"),
              detail: getLocalizedString(
                "core.createProjectQuestion.llmServiceOpenAIOption.detail"
              ),
            },
          ],
          skipSingleOption: true,
          default: "llm-service-azure-openai",
        },
        children: [
          {
            condition: { equals: "llm-service-azure-openai" },
            data: {
              type: "text",
              password: true,
              name: QuestionNames.AzureOpenAIKey,
              title: getLocalizedString(
                "core.createProjectQuestion.llmService.azureOpenAIKey.title"
              ),
              placeholder: getLocalizedString(
                "core.createProjectQuestion.llmService.azureOpenAIKey.placeholder"
              ),
            },
            children: [
              {
                condition: (inputs: Inputs) => {
                  return inputs[QuestionNames.AzureOpenAIKey]?.length > 0;
                },
                data: {
                  type: "text",
                  name: QuestionNames.AzureOpenAIEndpoint,
                  title: getLocalizedString(
                    "core.createProjectQuestion.llmService.azureOpenAIEndpoint.title"
                  ),
                  placeholder: getLocalizedString(
                    "core.createProjectQuestion.llmService.azureOpenAIEndpoint.placeholder"
                  ),
                },
                children: [
                  {
                    condition: (inputs: Inputs) => {
                      return inputs[QuestionNames.AzureOpenAIEndpoint]?.length > 0;
                    },
                    data: {
                      type: "text",
                      name: QuestionNames.AzureOpenAIDeploymentName,
                      title: getLocalizedString(
                        "core.createProjectQuestion.llmService.azureOpenAIDeploymentName.title"
                      ),
                      placeholder: getLocalizedString(
                        "core.createProjectQuestion.llmService.azureOpenAIDeploymentName.placeholder"
                      ),
                    },
                  },
                ],
              },
            ],
          },
          {
            condition: { equals: "llm-service-openai" },
            data: {
              type: "text",
              password: true,
              name: QuestionNames.OpenAIKey,
              title: getLocalizedString("core.createProjectQuestion.llmService.openAIKey.title"),
              placeholder: getLocalizedString(
                "core.createProjectQuestion.llmService.openAIKey.placeholder"
              ),
            },
          },
        ],
      },
    ],
  };
}

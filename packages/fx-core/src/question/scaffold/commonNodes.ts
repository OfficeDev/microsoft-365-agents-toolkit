// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ConditionFunc,
  Inputs,
  IQTreeNode,
  StringArrayValidation,
  StringValidation,
} from "@microsoft/teamsfx-api";
import { featureFlagManager, FeatureFlags } from "../../common/featureFlags";
import { getLocalizedString } from "../../common/localizeUtils";
import {
  apiOperationQuestion,
  apiSpecLocationQuestion,
  apiSpecTypeSelectQuestion,
  searchOpenAPISpecQueryQuestion,
  selectOpenApiSpecQuestion,
} from "../create";
import { QuestionNames } from "../questionNames";

export function azureOpenAINode(
  condition?: StringValidation | StringArrayValidation | ConditionFunc
): IQTreeNode {
  return {
    condition: condition,
    data: {
      type: "text",
      password: true,
      name: QuestionNames.AzureOpenAIKey,
      title: getLocalizedString("core.createProjectQuestion.llmService.azureOpenAIKey.title"),
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
  };
}

export function llmServiceNode(
  condition?: StringValidation | StringArrayValidation | ConditionFunc
): IQTreeNode {
  return {
    condition: condition,
    data: {
      type: "singleSelect",
      name: QuestionNames.LLMService,
      title: getLocalizedString("core.createProjectQuestion.llmService.title"),
      placeholder: getLocalizedString("core.createProjectQuestion.llmService.placeholder"),
      staticOptions: [
        {
          id: "llm-service-azure-openai",
          label: getLocalizedString("core.createProjectQuestion.llmServiceAzureOpenAIOption.label"),
          detail: getLocalizedString(
            "core.createProjectQuestion.llmServiceAzureOpenAIOption.detail"
          ),
        },
        {
          id: "llm-service-openai",
          label: getLocalizedString("core.createProjectQuestion.llmServiceOpenAIOption.label"),
          detail: getLocalizedString("core.createProjectQuestion.llmServiceOpenAIOption.detail"),
        },
      ],
      skipSingleOption: true,
      default: "llm-service-azure-openai",
    },
    children: [
      azureOpenAINode({ equals: "llm-service-azure-openai" }),
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
  };
}

export function apiSpecNode(condition: StringValidation | ConditionFunc): IQTreeNode {
  return {
    condition: condition,
    data: { type: "group", name: QuestionNames.FromExistingApi },
    children: [
      {
        data: apiSpecLocationQuestion(),
      },
      {
        condition: (inputs: Inputs) => {
          return !inputs[QuestionNames.ActionManifestPath];
        },
        data: apiOperationQuestion(),
      },
    ],
  };
}

export function apiSpecWithSearchNode(): IQTreeNode {
  return {
    data: { type: "group", name: QuestionNames.FromExistingApi },
    condition: { equals: "api-spec" },
    children: [inputOrSearchAPISpecNode()],
  };
}

export function inputOrSearchAPISpecNode(): IQTreeNode {
  return {
    data: apiSpecTypeSelectQuestion(),
    condition: (inputs: Inputs) => {
      return featureFlagManager.getBooleanValue(FeatureFlags.KiotaNPMIntegration);
    },
    children: [
      {
        condition: { equals: "enter-url-or-open-local-file" },
        data: apiSpecLocationQuestion(),
        children: [
          {
            condition: (inputs: Inputs) => {
              return !inputs[QuestionNames.ActionManifestPath];
            },
            data: apiOperationQuestion(true, true),
          },
        ],
      },
      {
        condition: { equals: "search-api" },
        data: searchOpenAPISpecQueryQuestion(),
        children: [
          {
            data: selectOpenApiSpecQuestion(),
          },
          {
            condition: (inputs: Inputs) => {
              return !!inputs[QuestionNames.SelectOpenApiSpec];
            },
            data: apiOperationQuestion(true, true),
          },
        ],
      },
    ],
  };
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { QuestionNames } from "../questionNames";
import { TemplateNames } from "../templateNames";

export const teamsNode = {
  condition: {
    equals: "teams-agent-and-app-type",
  },
  data: {
    title: "template.teams.title",
    name: QuestionNames.teamsAgentAndAppType,
    type: "singleSelect",
    options: [
      {
        id: TemplateNames.CustomCopilotBasic,
        label: "template.teams.general.label",
        detail: "template.teams.general.detail",
        data: TemplateNames.CustomCopilotBasic,
      },
      {
        id: QuestionNames.customCopilotRagType,
        label: "template.teams.rag.label",
        detail: "template.teams.rag.detail",
      },
      {
        id: TemplateNames.TeamsCollaboratorAgent,
        label: "template.teams.collaboratorAgent.label",
        detail: "template.teams.collaboratorAgent.detail",
        data: TemplateNames.TeamsCollaboratorAgent,
      },
      {
        id: QuestionNames.teamsOtherAppType,
        label: "template.teams.others.label",
        detail: "template.teams.others.detail",
      },
    ],
    placeholder: "template.customEngineAgent.placeholder",
  },
  children: [
    {
      condition: {
        equals: QuestionNames.customCopilotRagType,
      },
      data: {
        type: "singleSelect",
        name: QuestionNames.customCopilotRagType,
        title: "template.teams.rag.label",
        placeholder: "template.teams.rag.source.placeholder",
        default: TemplateNames.CustomCopilotRagCustomize,
        options: [
          {
            id: TemplateNames.CustomCopilotRagCustomize,
            label: "template.teams.rag.source.customize.label",
            detail: "template.teams.rag.source.customize.detail",
            data: TemplateNames.CustomCopilotRagCustomize,
          },
          {
            id: TemplateNames.CustomCopilotRagAzureAISearch,
            label: "template.teams.rag.source.azureAISearch.label",
            detail: "template.teams.rag.source.azureAISearch.detail",
            data: TemplateNames.CustomCopilotRagAzureAISearch,
          },
          {
            id: TemplateNames.CustomCopilotRagCustomApi,
            label: "template.teams.rag.source.customApi.label",
            detail: "template.teams.rag.source.customApi.detail",
            data: TemplateNames.CustomCopilotRagCustomApi,
          },
        ],
      },
      children: [
        {
          condition: {
            equals: TemplateNames.CustomCopilotRagCustomApi,
          },
          node: "apiSpecNode",
        },
      ],
    },
    {
      condition: {
        enum: [TemplateNames.CustomCopilotBasic, QuestionNames.customCopilotRagType],
      },
      node: "llmServiceNode",
    },
    {
      condition: {
        equals: TemplateNames.TeamsCollaboratorAgent,
      },
      node: "azureOpenAINode",
    },
    {
      condition: {
        equals: QuestionNames.teamsOtherAppType,
      },
      data: {
        type: "singleSelect",
        name: QuestionNames.teamsOtherAppType,
        title: "template.teams.others.capability.title",
        options: [
          {
            id: TemplateNames.Tab,
            label: "template.teams.others.tab.label",
            detail: "template.teams.others.tab.detail",
            data: TemplateNames.Tab,
          },
          {
            id: TemplateNames.DefaultMessageExtension,
            label: "template.teams.others.messageExtension.label",
            detail: "template.teams.others.messageExtension.detail",
            data: TemplateNames.DefaultMessageExtension,
          },
          {
            id: TemplateNames.DefaultBot,
            label: "template.teams.others.bot.label",
            detail: "template.teams.others.bot.detail",
            data: TemplateNames.DefaultBot,
          },
        ],
        placeholder: "template.customEngineAgent.placeholder",
      },
      children: [],
    },
  ],
};

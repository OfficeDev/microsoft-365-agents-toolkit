import { QuestionNames } from "../questionNames";
import { TemplateNames } from "../templateName";

export const teamsNode = {
  condition: {
    equals: "teams-agent-and-app-type",
  },
  data: {
    title: "core.createProjectQuestion.projectType.teamsAgentsAndApps.title",
    name: QuestionNames.teamsAgentAndAppType,
    type: "singleSelect",
    options: [
      {
        id: TemplateNames.CustomCopilotBasic,
        label: "core.createProjectQuestion.capability.customCopilotBasicOption.label",
        detail: "core.createProjectQuestion.capability.customCopilotBasicOption.detail",
        data: TemplateNames.CustomCopilotBasic,
      },
      {
        id: QuestionNames.customCopilotRagType,
        label: "core.createProjectQuestion.capability.customCopilotRagOption.label",
        detail: "core.createProjectQuestion.capability.customCopilotRagOption.detail",
      },
      {
        id: TemplateNames.TeamsCollaboratorAgent,
        label: "core.createProjectQuestion.capability.teamsAgent.collaborator.label",
        detail: "core.createProjectQuestion.capability.teamsAgent.collaborator.detail",
        data: TemplateNames.TeamsCollaboratorAgent,
      },
      {
        id: QuestionNames.teamsOtherAppType,
        label: "core.createProjectQuestion.capability.teamsAgent.others.label",
        detail: "core.createProjectQuestion.capability.teamsAgent.others.detail",
      },
    ],
    placeholder: "question.placeholder.choose",
  },
  children: [
    {
      condition: {
        equals: QuestionNames.customCopilotRagType,
      },
      data: {
        type: "singleSelect",
        name: QuestionNames.customCopilotRagType,
        title: "core.createProjectQuestion.capability.customCopilotRagOption.label",
        placeholder: "core.createProjectQuestion.capability.customCopilotRag.placeholder",
        default: TemplateNames.CustomCopilotRagCustomize,
        options: [
          {
            id: TemplateNames.CustomCopilotRagCustomize,
            label: "core.createProjectQuestion.capability.customCopilotRagCustomizeOption.label",
            detail: "core.createProjectQuestion.capability.customCopilotRagCustomizeOption.detail",
            data: TemplateNames.CustomCopilotRagCustomize,
          },
          {
            id: TemplateNames.CustomCopilotRagAzureAISearch,
            label:
              "core.createProjectQuestion.capability.customCopilotRagAzureAISearchOption.label",
            detail:
              "core.createProjectQuestion.capability.customCopilotRagAzureAISearchOption.detail",
            data: TemplateNames.CustomCopilotRagAzureAISearch,
          },
          {
            id: TemplateNames.CustomCopilotRagCustomApi,
            label: "core.createProjectQuestion.capability.customCopilotRagCustomApiOption.label",
            detail: "core.createProjectQuestion.capability.customCopilotRagCustomApiOption.detail",
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
        title: "core.createProjectQuestion.teamsCapability.title",
        options: [
          {
            id: TemplateNames.Tab,
            label: "core.TabNonSso.label",
            detail: "core.TabNonSso.detail",
            data: TemplateNames.Tab,
          },
          {
            id: TemplateNames.DefaultMessageExtension,
            label: "core.MessageExtensionOption.label",
            detail: "core.MessageExtensionOption.detail",
            data: TemplateNames.DefaultMessageExtension,
          },
          {
            id: TemplateNames.DefaultBot,
            label: "core.BotNewUIOption.label",
            detail: "core.BotNewUIOption.detail",
            data: TemplateNames.DefaultBot,
          },
        ],
        placeholder: "question.placeholder.choose",
      },
      children: [],
    },
  ],
};

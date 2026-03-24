// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * TDP (Teams Developer Portal) project type node.
 * Shown when TDP can't auto-determine the template type.
 * Subset of root wizard options — only DA, CEA, Teams supported.
 */
export const tdpNode = {
  data: {
    title: "core.createProjectQuestion.title",
    name: "project-type",
    type: "singleSelect",
    options: [
      {
        id: "copilot-agent-type",
        label: "core.createProjectQuestion.projectType.declarativeAgent.label",
        detail: "core.createProjectQuestion.projectType.declarativeAgent.detail",
        groupName: "core.createProjectQuestion.projectType.createGroup.aiAgent",
        icon: "$(teamsfx-agent)",
      },
      {
        id: "custom-engine-agent-type",
        label: "core.createProjectQuestion.projectType.customCopilot.label",
        detail: "core.createProjectQuestion.projectType.customCopilot.detail",
        groupName: "core.createProjectQuestion.projectType.createGroup.aiAgent",
        icon: "$(teamsfx-custom-copilot)",
      },
      {
        id: "teams-agent-and-app-type",
        label: "core.createProjectQuestion.projectType.teamsAgentsAndApps.label",
        detail: "core.createProjectQuestion.projectType.teamsAgentsAndApps.detail",
        groupName: "core.createProjectQuestion.projectType.createGroup.m365Apps",
        icon: "$(microsoft365-agents-toolkit-teams)",
      },
    ],
  },
  children: [
    { node: "daNode", condition: { equals: "copilot-agent-type" } },
    { node: "ceaNode", condition: { equals: "custom-engine-agent-type" } },
    { node: "teamsNode", condition: { equals: "teams-agent-and-app-type" } },
  ],
};

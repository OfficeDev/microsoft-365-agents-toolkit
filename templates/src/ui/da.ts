// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Declarative Agent sub-tree for the wizard.
 * Serialized to daNode.json at build time.
 */
export const daNode = {
  data: {
    title: "core.createProjectQuestion.projectType.copilotExtension.title",
    placeholder: "core.createProjectQuestion.projectType.copilotExtension.placeholder",
    name: "capabilities",
    type: "singleSelect",
    skipSingleOption: true,
    options: [
      {
        id: "declarative-agent",
        label: "core.createProjectQuestion.projectType.declarativeAgent.label",
        detail: "core.createProjectQuestion.projectType.declarativeAgent.detail",
      },
    ],
  },
  children: [
    {
      condition: { equals: "declarative-agent" },
      data: {
        name: "with-plugin",
        title: "core.createProjectQuestion.declarativeCopilot.title",
        type: "singleSelect",
        placeholder: "core.createProjectQuestion.declarativeCopilot.placeholder",
        options: [
          {
            id: "no",
            label: "core.createProjectQuestion.noPlugin.label",
            detail: "core.createProjectQuestion.noPlugin.detail",
            data: "copilot-gpt-basic",
          },
          {
            id: "yes",
            label: "core.createProjectQuestion.addPlugin.label",
            detail: "core.createProjectQuestion.addPlugin.detail",
          },
          {
            id: "gc",
            label: "core.createProjectQuestion.addGC.label",
            detail: "core.createProjectQuestion.addGC.detail",
            data: "declarative-agent-with-graph-connector",
          },
          {
            id: "type-spec",
            label: "core.createProjectQuestion.apiPlugin.typeSpec.label",
            detail: "core.createProjectQuestion.apiPlugin.typeSpec.detail",
            data: "declarative-agent-typespec",
          },
        ],
      },
      children: [
        {
          condition: { equals: "yes" },
          data: {
            name: "action-type",
            title: "core.createProjectQuestion.createApiPlugin.title",
            placeholder: "core.createProjectQuestion.addApiPlugin.placeholder",
            type: "singleSelect",
            options: [
              {
                id: "new-api",
                label: "core.createProjectQuestion.capability.copilotPluginNewApiOption.label",
                detail: "core.createProjectQuestion.capability.copilotPluginNewApiOption.detail",
              },
              {
                id: "api-spec",
                label: "core.createProjectQuestion.capability.copilotPluginApiSpecOption.label",
                detail: "core.createProjectQuestion.capability.copilotPluginApiSpecOption.detail",
                data: "api-plugin-from-existing-api",
              },
              {
                id: "da-meta-os",
                label: "core.createProjectQuestion.capability.DAMetaOS.label",
                detail: "core.createProjectQuestion.capability.DAMetaOS.detail",
                data: "declarative-agent-meta-os-new-project",
                featureFlag: "TEAMSFX_DA_METAOS",
              },
              {
                id: "mcp",
                label: "core.createProjectQuestion.mcpForDa.label",
                detail: "core.createProjectQuestion.mcpForDa.detail",
                data: "declarative-agent-with-action-from-mcp",
                featureFlag: "TEAMSFX_MCP_FOR_DA",
              },
            ],
          },
          children: [
            {
              condition: { equals: "new-api" },
              data: {
                name: "api-auth",
                title: "core.createProjectQuestion.apiMessageExtensionAuth.title",
                placeholder: "core.createProjectQuestion.apiMessageExtensionAuth.placeholder",
                type: "singleSelect",
                options: [
                  { id: "none", label: "None", data: "api-plugin-from-scratch" },
                  {
                    id: "api-key",
                    label: "API Key",
                    data: "api-plugin-from-scratch-bearer",
                  },
                  {
                    id: "microsoft-entra",
                    label: "Microsoft Entra",
                    data: "api-plugin-from-scratch-oauth",
                  },
                  {
                    id: "oauth",
                    label: "OAuth",
                    data: "api-plugin-from-scratch-oauth",
                  },
                ],
              },
            },
            { node: "apiSpecNode", condition: { equals: "api-spec" } },
            { node: "mcpServerTypeNode", condition: { equals: "mcp" } },
          ],
        },
      ],
    },
  ],
};

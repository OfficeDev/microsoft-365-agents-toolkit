// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Add Action tree for the "Add Action to Declarative Agent" wizard.
 * Build-time: serialized to addActionNode.json
 * Runtime: loaded by fx-core and resolved via constructNode()
 *
 * Node references (e.g. { node: "apiPluginStartNode" }) are resolved at runtime
 * by constructNode.ts to TypeScript-defined sub-trees with dynamic behavior.
 */
export const addActionNode = {
  data: {
    type: "singleSelect",
    name: "action-type",
    title: "core.createProjectQuestion.addApiPlugin.title",
    placeholder: "template.createProjectQuestion.addApiPlugin.placeholder",
    options: [
      {
        id: "api-spec",
        label: "template.createProjectQuestion.capability.copilotPluginApiSpecOption.label",
        detail: "template.createProjectQuestion.capability.copilotPluginApiSpecOption.detail",
      },
    ],
  },
  children: [
    { node: "apiSpecWithSearchNode", condition: { equals: "api-spec" } },
    { node: "selectTeamsAppManifestNode" },
  ],
};

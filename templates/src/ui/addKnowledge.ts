// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Add Knowledge tree for the "Add Knowledge to Declarative Agent" wizard.
 * Build-time: serialized to addKnowledgeNode.json
 * Runtime: loaded by fx-core and resolved via constructNode()
 *
 * Node references (e.g. { node: "addKnowledgeStartNode" }) are resolved at runtime
 * by constructNode.ts to TypeScript-defined sub-trees with dynamic behavior.
 */
export const addKnowledgeNode = {
  data: {
    type: "singleSelect",
    name: "knowledge-source",
    title: "core.createProjectQuestion.addKnowledge.title",
    placeholder: "core.createProjectQuestion.addKnowledge.placeholder",
    options: [
      {
        id: "web-search",
        label: "core.createProjectQuestion.capability.knowledgeWebSearch.label",
        detail: "core.createProjectQuestion.capability.knowledgeWebSearch.detail",
      },
      {
        id: "oneDrive-sharePoint",
        label: "core.createProjectQuestion.capability.knowledgeOneDriveSharePoint.label",
        detail: "core.createProjectQuestion.capability.knowledgeOneDriveSharePoint.detail",
      },
      {
        id: "graph-connector",
        label: "core.createProjectQuestion.capability.knowledgeGraphConnector.label",
        detail: "core.createProjectQuestion.capability.knowledgeGraphConnector.detail",
      },
      {
        id: "embedded-knowledge",
        label: "core.createProjectQuestion.capability.knowledgeEmbeddedKnowledge.label",
        detail: "core.createProjectQuestion.capability.knowledgeEmbeddedKnowledge.detail",
      },
    ],
  },
  children: [
    {
      node: "searchTypeNode",
      condition: { equals: "web-search" },
      children: [
        { node: "webContentNode", condition: { equals: "url" } },
        { node: "selectTeamsAppManifestNode" },
      ],
    },
    {
      node: "searchTypeNode",
      condition: { equals: "oneDrive-sharePoint" },
      children: [
        { node: "oneDriveSharePointItemNode", condition: { equals: "url" } },
        { node: "oneDriveSharePointItemConfirmNode", condition: { equals: "url" } },
        { node: "selectTeamsAppManifestNode" },
      ],
    },
    {
      node: "gcItemNode",
      condition: { equals: "graph-connector" },
      children: [
        { node: "gcListNode", condition: { equals: "listConnections" } },
        { node: "gcInputNode", condition: { equals: "inputConnectionId" } },
        { node: "selectTeamsAppManifestNode" },
      ],
    },
    {
      node: "selectTeamsAppManifestNode",
      condition: { equals: "embedded-knowledge" },
      children: [
        { node: "embeddedKnowledgeNode" },
      ],
    },
  ],
};

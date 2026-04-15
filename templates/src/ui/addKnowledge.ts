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
  node: "addKnowledgeStartNode",
  children: [
    { node: "addKnowledgeWebSearchNode", condition: { equals: "web-search" } },
    { node: "addKnowledgeOneDriveNode", condition: { equals: "oneDrive-sharePoint" } },
    { node: "addKnowledgeGCNode", condition: { equals: "graph-connector" } },
    { node: "addKnowledgeEmbeddedNode", condition: { equals: "embedded-knowledge" } },
  ],
};

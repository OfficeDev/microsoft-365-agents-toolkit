// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Graph Connector sub-tree for the wizard.
 * Serialized to graphConnectorNode.json at build time.
 * This is a simple group node — the project is created directly
 * after folder/name selection.
 */
export const graphConnectorNode = {
  data: {
    type: "group",
  },
  children: [{ node: "gcNameNode" }, { node: "gcConnectionIdNode" }],
};

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
  node: "apiPluginStartNode",
  children: [
    { node: "addActionApiSpecNode", condition: { equals: "api-spec" } },
    { node: "mcpForDANode", condition: { equals: "mcp" } },
    { node: "selectTeamsAppManifestNode" },
  ],
};

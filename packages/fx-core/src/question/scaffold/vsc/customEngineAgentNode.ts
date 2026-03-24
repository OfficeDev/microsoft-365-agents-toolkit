// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { IQTreeNode, Platform } from "@microsoft/teamsfx-api";
import { getRootProjectTypeNode } from "./rootNode";

/**
 * Extract the Custom Engine Agent sub-tree from the combined wizardNode.json.
 * Used by TDP (Teams Developer Portal) flow.
 */
export function getCustomEngineAgentNode(): IQTreeNode {
  const root = getRootProjectTypeNode(Platform.VSCode);
  const ceaNode = root.children?.find(
    (c) => (c.condition as any)?.equals === "custom-engine-agent-type"
  );
  return ceaNode ?? { data: { type: "group" } };
}

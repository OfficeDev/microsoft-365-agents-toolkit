// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Combined wizard node for the "Create New Agent/App" wizard.
 * Build-time: this is serialized to wizardNode.json
 * Runtime: loaded from ~/.fx/ui/wizardNode.json (cache) or bundled fallback
 *
 * Source sub-trees are defined separately for maintainability
 * but merged here into a single JSON for distribution.
 */
import { ceaNode } from "./cea";
import { daNode } from "./da";
import { graphConnectorNode } from "./graphConnector";
import { officeAddinNode } from "./officeAddin";
import { rootNode } from "./root";
import { teamsNode } from "./teams";

// Deep clone and inline all sub-tree references into the root node
function buildWizardNode() {
  const wizard = JSON.parse(JSON.stringify(rootNode));

  // Map node references to their actual sub-tree definitions
  const nodeMap: Record<string, Record<string, unknown>> = {
    daNode: daNode,
    ceaNode: ceaNode,
    graphConnectorNode: graphConnectorNode,
    teamsNode: teamsNode,
    officeAddinNode: officeAddinNode,
  };

  // Replace node references with inlined sub-trees
  wizard.children = wizard.children.map((child: any) => {
    if (child.node && nodeMap[child.node]) {
      const inlined = JSON.parse(JSON.stringify(nodeMap[child.node]));
      // Merge condition from the reference into the inlined node
      if (child.condition) {
        inlined.condition = child.condition;
      }
      return inlined;
    }
    return child;
  });

  return wizard;
}

export const wizardNode = buildWizardNode();

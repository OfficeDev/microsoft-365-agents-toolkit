// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Composite runtime nodes for the "Add Knowledge" flow.
 * These encapsulate complex branching logic (dynamic options, validation)
 * that cannot be expressed in static JSON.
 *
 * Referenced by constructNode.ts via lazy require() to avoid circular dependencies.
 */

import { IQTreeNode } from "@microsoft/teamsfx-api";
import { GCSelectOptions, KnowledgeSearchTypeOptions } from "../constants";
import {
  GCInputQuestion,
  GCItemQuestion,
  GCListQuestion,
  oneDriveSharePointItemConfirmQuestion,
  oneDriveSharePointItemQuestion,
  searchTypeQuestion,
  webContentQuestion,
} from "../create";
import { addEmbeddedKnowledgeFilesQuestion, selectTeamsAppManifestQuestion } from "../other";

/**
 * Composite node: Web Search branch for Add Knowledge.
 * searchType → webContent (if URL) → selectManifest
 */
export function addKnowledgeWebSearchNode(): IQTreeNode {
  return {
    data: searchTypeQuestion(),
    children: [
      {
        data: webContentQuestion(),
        condition: { equals: KnowledgeSearchTypeOptions.url().id },
      },
      {
        data: selectTeamsAppManifestQuestion(),
      },
    ],
  };
}

/**
 * Composite node: OneDrive/SharePoint branch for Add Knowledge.
 * searchType → oneDriveSharePointItem (if URL) → confirm (if URL) → selectManifest
 */
export function addKnowledgeOneDriveNode(): IQTreeNode {
  return {
    data: searchTypeQuestion(),
    children: [
      {
        data: oneDriveSharePointItemQuestion(),
        condition: { equals: KnowledgeSearchTypeOptions.url().id },
      },
      {
        data: oneDriveSharePointItemConfirmQuestion(),
        condition: { equals: KnowledgeSearchTypeOptions.url().id },
      },
      {
        data: selectTeamsAppManifestQuestion(),
      },
    ],
  };
}

/**
 * Composite node: Graph Connector (Copilot connector) branch for Add Knowledge.
 * GCItem → GCList (if list) / GCInput (if input) → selectManifest
 */
export function addKnowledgeGCNode(): IQTreeNode {
  return {
    data: GCItemQuestion(),
    children: [
      {
        data: GCListQuestion(),
        condition: { equals: GCSelectOptions.list().id },
      },
      {
        data: GCInputQuestion(),
        condition: { equals: GCSelectOptions.input().id },
      },
      {
        data: selectTeamsAppManifestQuestion(),
      },
    ],
  };
}

/**
 * Composite node: Embedded Knowledge branch for Add Knowledge.
 * selectManifest → embeddedKnowledgeFiles
 */
export function addKnowledgeEmbeddedNode(): IQTreeNode {
  return {
    data: selectTeamsAppManifestQuestion(),
    children: [
      {
        data: addEmbeddedKnowledgeFilesQuestion(),
      },
    ],
  };
}

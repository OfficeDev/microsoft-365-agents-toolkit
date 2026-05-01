// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 * Import `DeclarativeAgentManifest` or version-specific types like `DeclarativeAgentManifestV1D6` from the generated-types module.
 */

import { ConversationStarter } from "./pluginManifest";

/**
 * @deprecated Use capabilities from generated types instead.
 */
export enum DeclarativeCopilotCapabilityName {
  WebSearch = "WebSearch",
  GraphicArt = "GraphicArt",
  CodeInterpreter = "CodeInterpreter",
  OneDriveAndSharePoint = "OneDriveAndSharePoint",
  GraphConnectors = "GraphConnectors",
  EmbeddedKnowledge = "EmbeddedKnowledge",
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface WebSearchCapability {
  name: DeclarativeCopilotCapabilityName.WebSearch;
  sites?: Site[];
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface GraphicArtCapability {
  name: DeclarativeCopilotCapabilityName.GraphicArt;
  [k: string]: unknown;
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface CodeInterpreterCapability {
  name: DeclarativeCopilotCapabilityName.CodeInterpreter;
  [k: string]: unknown;
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface OneDriveAndSharePointCapability {
  name: DeclarativeCopilotCapabilityName.OneDriveAndSharePoint;
  items_by_sharepoint_ids?: SharePointIDs[];
  items_by_url?: Site[];
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface GraphConnectorsCapability {
  name: DeclarativeCopilotCapabilityName.GraphConnectors;
  connections: Connection[];
  [k: string]: unknown;
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface EmbeddedKnowledgeCapability {
  name: DeclarativeCopilotCapabilityName.EmbeddedKnowledge;
  files: File[];
}

/**
 * @deprecated Use `DeclarativeAgentManifest` or version-specific types from `./generated-types` instead.
 */
export interface DeclarativeCopilotManifestSchema {
  id?: string;
  name: string;
  description: string;
  instructions?: string;
  capabilities?: (
    | WebSearchCapability
    | GraphicArtCapability
    | CodeInterpreterCapability
    | OneDriveAndSharePointCapability
    | GraphConnectorsCapability
    | EmbeddedKnowledgeCapability
  )[];
  conversation_starters?: ConversationStarter[];
  actions?: ActionObject[];
  agent_skills?: AgentSkillObject[];
  [k: string]: unknown;
}

/**
 * @deprecated Use `AgentSkillElement` from `./generated-types` instead.
 */
export interface AgentSkillObject {
  folder: string;
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface SharePointIDs {
  site_id?: string;
  web_id?: string;
  list_id?: string;
  unique_id?: string;
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface File {
  file_name?: string;
  file?: string;
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface Site {
  url: string;
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface Connection {
  connection_id: string;
  [k: string]: unknown;
}

/**
 * @deprecated Use auto-generated types from `./generated-types` instead.
 */
export interface ActionObject {
  id: string;
  file: string;
  [k: string]: unknown;
}

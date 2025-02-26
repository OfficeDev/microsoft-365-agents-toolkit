// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ConversationStarter } from "./pluginManifest";

export enum DeclarativeCopilotCapabilityName {
  WebSearch = "WebSearch",
  GraphicArt = "GraphicArt",
  CodeInterpreter = "CodeInterpreter",
  OneDriveAndSharePoint = "OneDriveAndSharePoint",
  GraphConnectors = "GraphConnectors",
  EmbeddedKnowledge = "EmbeddedKnowledge",
}

export interface WebSearchCapability {
  name: DeclarativeCopilotCapabilityName.WebSearch;
  sites?: Site[];
}

export interface GraphicArtCapability {
  name: DeclarativeCopilotCapabilityName.GraphicArt;
  [k: string]: unknown;
}

export interface CodeInterpreterCapability {
  name: DeclarativeCopilotCapabilityName.CodeInterpreter;
  [k: string]: unknown;
}

export interface OneDriveAndSharePointCapability {
  name: DeclarativeCopilotCapabilityName.OneDriveAndSharePoint;
  items_by_sharepoint_ids?: File[];
  items_by_url?: Site[];
}

export interface GraphConnectorsCapability {
  name: DeclarativeCopilotCapabilityName.GraphConnectors;
  connections: Connection[];
  [k: string]: unknown;
}

export interface EmbeddedKnowledgeCapability {
  name: DeclarativeCopilotCapabilityName.EmbeddedKnowledge;
  files: File[];
}

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
  [k: string]: unknown;
}

export interface File {
  site_id?: string;
  web_id?: string;
  list_id?: string;
  unique_id?: string;
  file_name?: string;
  file?: string;
}

export interface Site {
  url: string;
}

export interface Connection {
  connection_id: string;
  [k: string]: unknown;
}

export interface ActionObject {
  id: string;
  file: string;
  [k: string]: unknown;
}

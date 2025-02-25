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

export interface DeclarativeCopilotManifestSchema {
  id?: string;
  name: string;
  description: string;
  instructions?: string;
  capabilities?: (
    | {
        name: DeclarativeCopilotCapabilityName.WebSearch;
        [k: string]: unknown;
      }
    | {
        name: DeclarativeCopilotCapabilityName.GraphicArt;
        [k: string]: unknown;
      }
    | {
        name: DeclarativeCopilotCapabilityName.CodeInterpreter;
        [k: string]: unknown;
      }
    | {
        name: DeclarativeCopilotCapabilityName.OneDriveAndSharePoint;
        items_by_sharepoint_ids?: File[];
        items_by_url?: Site[];
      }
    | {
        name: DeclarativeCopilotCapabilityName.GraphConnectors;
        connections: Connection[];
        [k: string]: unknown;
      }
    | {
        name: DeclarativeCopilotCapabilityName.EmbeddedKnowledge;
        files: File[];
      }
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

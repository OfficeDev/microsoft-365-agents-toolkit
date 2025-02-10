// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ConversationStarter } from "./pluginManifest";

export interface DeclarativeCopilotManifestSchema {
  id?: string;
  name: string;
  description: string;
  instructions?: string;
  capabilities?: (
    | {
        name: "WebSearch";
        [k: string]: unknown;
      }
    | {
        name: "GraphicArt";
        [k: string]: unknown;
      }
    | {
        name: "CodeInterpreter";
        [k: string]: unknown;
      }
    | {
        name: "OneDriveAndSharePoint";
        items_by_sharepoint_ids?: File[];
        items_by_url?: Site[];
      }
    | {
        name: "GraphConnectors";
        connections: Connection[];
        [k: string]: unknown;
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

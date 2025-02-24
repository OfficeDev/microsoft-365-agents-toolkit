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
        name: "SharePoint";
        files?: File[];
        sites?: Site[];
        [k: string]: unknown;
      }
    | {
        name: "OneDrive";
        files: File[];
        [k: string]: unknown;
      }
    | {
        name: "GraphConnectors";
        connections: Connection[];
        [k: string]: unknown;
      }
    | {
        name: "EmbeddedKnowledge";
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
  path: string;
  site_name: string;
  [k: string]: unknown;
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

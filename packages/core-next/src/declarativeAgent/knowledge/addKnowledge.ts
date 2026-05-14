// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result, err } from "neverthrow";
import type { AtkError } from "../../core/error";
import { userError } from "../../core/error";
import type { KnowledgeSource, ODSPItemMetadata } from "../types";
import { addWebSearchKnowledge } from "./webSearch";
import { addOneDriveSharePointKnowledge } from "./oneDriveSharePoint";
import { addGraphConnectorKnowledge } from "./graphConnector";
import { addEmbeddedKnowledge } from "./embeddedKnowledge";

/**
 * Input for the add-knowledge dispatcher.
 */
export interface AddKnowledgeInput {
  agentManifestPath: string;
  source: KnowledgeSource;

  /** (web-search) Optional site URL to scope the search. */
  siteUrl?: string;

  /** (onedrive-sharepoint) Optional item metadata. */
  odspItem?: ODSPItemMetadata;

  /** (graph-connector) One or more connection IDs. */
  connectionIds?: string[];

  /** (embedded-knowledge) Source file paths to embed. */
  embeddedFilePaths?: string[];
}

/**
 * Dispatch to the appropriate knowledge handler based on source type.
 *
 * This is the single entry-point that operations.ts calls.
 */
export async function addKnowledge(input: AddKnowledgeInput): Promise<Result<void, AtkError>> {
  switch (input.source) {
    case "web-search":
      return addWebSearchKnowledge(input.agentManifestPath, input.siteUrl);

    case "onedrive-sharepoint":
      return addOneDriveSharePointKnowledge(input.agentManifestPath, input.odspItem);

    case "graph-connector":
      if (!input.connectionIds?.length) {
        return err(
          userError(
            "MissingConnectionIds",
            "At least one Graph Connector connection ID is required.",
            { source: "declarativeAgent/knowledge" }
          )
        );
      }
      return addGraphConnectorKnowledge(input.agentManifestPath, input.connectionIds);

    case "embedded-knowledge":
      if (!input.embeddedFilePaths?.length) {
        return err(
          userError(
            "MissingEmbeddedFiles",
            "At least one file path is required for embedded knowledge.",
            { source: "declarativeAgent/knowledge" }
          )
        );
      }
      return addEmbeddedKnowledge(input.agentManifestPath, input.embeddedFilePaths);

    default: {
      const exhaustive: never = input.source;
      return err(
        userError("UnknownKnowledgeSource", `Unknown knowledge source: ${exhaustive}`, {
          source: "declarativeAgent/knowledge",
        })
      );
    }
  }
}

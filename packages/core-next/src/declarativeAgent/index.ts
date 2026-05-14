// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Declarative Agent module.
 *
 * Self-contained module that encapsulates all DA-related capabilities:
 * - Knowledge sources (web search, OneDrive/SharePoint, Graph Connectors, embedded)
 * - Actions (existing plugin, MCP-backed, remove)
 * - Auth injection (OAuth, API key, MCP OAuth)
 * - Capabilities (sensitivity labels, conversation starters)
 * - Manifest resolution (${{VAR}} placeholders)
 * - Operations (wired into the Operation pipeline with Zod validation)
 */

// === Types ===
export type {
  KnowledgeSource,
  ODSPItemMetadata,
  GraphConnectorItem,
  AuthScheme,
  AddExistingPluginInput,
  AddExistingPluginResult,
  AddMCPActionInput,
  MCPToolDefinition,
  AuthActionInjectResult,
  SetSensitivityLabelInput,
  EnvPlaceholder,
} from "./types";

export {
  DeclarativeAgentManifestWrapper,
  CapabilityName,
  APIPluginManifestWrapper,
  RuntimeType,
} from "./types";

// === Knowledge ===
export {
  addKnowledge,
  addWebSearchKnowledge,
  webSearchRequiresConfirmation,
  addOneDriveSharePointKnowledge,
  addGraphConnectorKnowledge,
  addEmbeddedKnowledge,
} from "./knowledge";
export type { AddKnowledgeInput, ODSPSharePointIds } from "./knowledge";

// === Actions ===
export { addExistingPlugin, addMCPAction, removeAction } from "./actions";

// === Auth ===
export {
  injectOAuthAction,
  injectApiKeyAction,
  injectMCPOAuthAction,
  findNextAvailableEnv,
} from "./auth";

// === Capabilities ===
export { setSensitivityLabel, setConversationStarters } from "./capabilities";

// === Manifest resolution ===
export { findPlaceholders, resolvePlaceholders, getAgentManifestPath } from "./manifest";

// === Operations (wired into the v4 Operation pipeline) ===
export {
  addKnowledgeOp,
  addExistingPluginOp,
  addMCPActionOp,
  removeActionOp,
  setSensitivityLabelOp,
  setConversationStartersOp,
  injectOAuthActionOp,
  injectApiKeyActionOp,
} from "./operations";

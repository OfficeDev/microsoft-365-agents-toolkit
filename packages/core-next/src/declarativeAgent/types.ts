// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * DA module types.
 *
 * Types from @microsoft/app-manifest are re-exported for convenience.
 * Module-specific types that are NOT covered by the manifest package live here.
 */

// Re-export manifest-package types used widely in the DA module
export type {
  DeclarativeAgentManifest,
  DeclarativeAgentManifestLatest,
  APIPluginManifest,
  APIPluginManifestLatest,
} from "@microsoft/app-manifest";

export {
  DeclarativeAgentManifestWrapper,
  CapabilityName,
  APIPluginManifestWrapper,
  RuntimeType,
} from "@microsoft/app-manifest";

export type { CapabilityNameValue, RuntimeTypeValue } from "@microsoft/app-manifest";

// ---------------------------------------------------------------------------
// Knowledge sources
// ---------------------------------------------------------------------------

/** Recognized knowledge source kinds. */
export type KnowledgeSource =
  | "web-search"
  | "onedrive-sharepoint"
  | "graph-connector"
  | "embedded-knowledge";

/** Metadata for an OneDrive / SharePoint item obtained from Graph API. */
export interface ODSPItemMetadata {
  id: string;
  name: string;
  uniqueId?: string;
  listId?: string;
  webId?: string;
  siteId?: string;
  webUrl?: string;
  itemType?: "File" | "Folder";
}

/** Graph Connector item (id + user-facing label). */
export interface GraphConnectorItem {
  id: string;
  label: string;
}

// ---------------------------------------------------------------------------
// Actions / plugins
// ---------------------------------------------------------------------------

/** Auth scheme recognized by the action subsystem. */
export type AuthScheme = "oauth" | "api-key" | "bearer-token" | "microsoft-entra" | "none";

/** Input for adding an action from an existing plugin manifest + spec. */
export interface AddExistingPluginInput {
  /** Absolute path to the declarative agent manifest (declarativeAgent.json). */
  agentManifestPath: string;
  /** Absolute path to the source plugin manifest (ai-plugin.json). */
  pluginManifestPath: string;
  /** Absolute path to the source OpenAPI spec. */
  apiSpecPath: string;
  /** Desired action ID in the agent manifest. */
  actionId: string;
}

/** Result returned after successfully adding an existing plugin. */
export interface AddExistingPluginResult {
  /** Warnings (e.g. unresolved env variables). */
  warnings: string[];
  /** Final destination path of the copied plugin manifest. */
  destinationPluginManifestPath: string;
}

/** Input for adding an action backed by an MCP server. */
export interface AddMCPActionInput {
  /** Absolute path to the plugin manifest (ai-plugin.json). */
  pluginManifestPath: string;
  /** MCP server URL (remote) or local identifier. */
  serverUrlOrIdentifier: string;
  /** Human-readable server name. */
  serverName: string;
  /** Whether the MCP server is local (via stdio) or remote (HTTP). */
  isLocal: boolean;
  /** Auth type for the MCP server. */
  auth: AuthScheme;
  /** OAuth metadata URLs (only when auth = "oauth"). */
  oauthMeta?: {
    authorizationUrl: string;
    tokenUrl: string;
    refreshUrl?: string;
  };
  /** All available tools advertised by the server. */
  availableTools: MCPToolDefinition[];
  /** Subset of tool names the user selected. */
  selectedTools: string[];
}

/** A single MCP tool definition. */
export interface MCPToolDefinition {
  name: string;
  description?: string;
  inputSchema?: Record<string, unknown>;
}

// ---------------------------------------------------------------------------
// Auth injection
// ---------------------------------------------------------------------------

/** Result of injecting an auth action into teamsapp.yml. */
export interface AuthActionInjectResult {
  /** Default env var name for the registration ID (e.g. OAUTH2_CONFIGURATION_ID). */
  defaultRegistrationIdEnvName: string | undefined;
  /** Actual env var name assigned (may have numeric suffix to avoid conflicts). */
  registrationIdEnvName: string | undefined;
}

// ---------------------------------------------------------------------------
// Sensitivity labels
// ---------------------------------------------------------------------------

/** Input for setting a sensitivity label on a DA manifest. */
export interface SetSensitivityLabelInput {
  /** Absolute path to the declarative agent manifest. */
  agentManifestPath: string;
  /** Sensitivity label ID to apply. */
  labelId: string;
}

// ---------------------------------------------------------------------------
// Env-variable resolution
// ---------------------------------------------------------------------------

/** A ${{VAR_NAME}} placeholder found in a manifest string. */
export interface EnvPlaceholder {
  /** Full token (e.g. "${{BOT_ID}}") */
  token: string;
  /** Variable name (e.g. "BOT_ID") */
  name: string;
}

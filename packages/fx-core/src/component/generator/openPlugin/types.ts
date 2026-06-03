// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface OpenPluginAuthorObject {
  name?: string;
  email?: string;
  url?: string;
}

export interface OpenPluginManifest {
  name: string;
  version?: string;
  description?: string;
  author?: string | OpenPluginAuthorObject;
  homepage?: string;
  repository?: string;
  license?: string;
  logo?: string;
  keywords?: string[];
  // Component-path overrides per Open Plugin Spec v1.0. Only string form is
  // supported in this converter; array/object forms throw.
  skills?: string | string[] | Record<string, unknown>;
  commands?: string | string[] | Record<string, unknown>;
  agents?: string | string[] | Record<string, unknown>;
  hooks?: string | string[] | Record<string, unknown>;
  mcpServers?: string | string[] | Record<string, unknown>;
  lspServers?: string | string[] | Record<string, unknown>;
  rules?: string | string[] | Record<string, unknown>;
  outputStyles?: string | string[];
}

export interface ParsedAuthor {
  name?: string;
  email?: string;
  url?: string;
}

export interface OpenPluginMcpServerEntry {
  url?: string;
  type?: string;
  description?: string;
  // Other fields tolerated but not used.
  [key: string]: unknown;
}

export interface OpenPluginMcpJson {
  mcpServers?: Record<string, OpenPluginMcpServerEntry>;
  // Tolerate the bare form where servers are at the root.
  [key: string]: unknown;
}

export interface ParsedOpenPlugin {
  pluginRoot: string;
  manifest: OpenPluginManifest;
  manifestPath: string;
  manifestKind: "open-plugin" | "claude-plugin" | "cursor-plugin";
  mcpServers: Record<string, OpenPluginMcpServerEntry>;
  skills: string[];
  skillsRoot?: string;
  commands: string[];
  commandsRoot?: string;
  hasColorPng: boolean;
  hasOutlinePng: boolean;
  warnings: string[];
  /** Round-trip metadata produced by `atk export openplugin`, when present. */
  atkExtension?: AtkExtensionBlock;
}

export type AuthorizationType = "None" | "OAuthPluginVault" | "ApiKeyPluginVault";
export type DefaultAuthOption = "Auto" | AuthorizationType;

/**
 * Extension block embedded under `x-microsoft-365-agents-toolkit` in plugin.json
 * by `atk export openplugin`. Carries every field that the Open Plugin Spec
 * cannot natively represent so that re-importing reconstructs the original
 * project losslessly. All fields are optional — the importer treats missing
 * keys the same way it treats a plugin.json without this extension at all.
 */
export interface AtkExtensionBlock {
  manifestVersion?: string;
  id?: string;
  packageName?: string;
  accentColor?: string;
  developer?: {
    name?: string;
    websiteUrl?: string;
    privacyUrl?: string;
    termsOfUseUrl?: string;
  };
  name?: { short?: string; full?: string };
  description?: { short?: string; full?: string };
  /**
   * Per-agentConnector overrides preserved verbatim: the keys are the
   * connector ids (matching the .mcp.json server name). Values store the
   * fields .mcp.json cannot carry: displayName, description, authorization.
   */
  agentConnectors?: Record<string, AtkAgentConnectorExt>;
}

export interface AtkAgentConnectorExt {
  displayName?: string;
  description?: string;
  authorization?: {
    type: AuthorizationType;
    referenceId?: string;
  };
}

export interface ImportInputs {
  path: string;
  output?: string;
  /** Optional when plugin.json carries an x-microsoft-365-agents-toolkit block. */
  privacyUrl?: string;
  /** Optional when plugin.json carries an x-microsoft-365-agents-toolkit block. */
  termsUrl?: string;
  websiteUrl?: string;
  appId?: string;
  defaultAuthType?: DefaultAuthOption;
  packageName?: string;
}

export interface ExportInputs {
  /** Path to the existing ATK project (folder that contains appPackage/manifest.json). */
  path: string;
  /** Destination open-plugin directory. Defaults to ./<plugin-name>-openplugin. */
  output?: string;
  /** Manifest kind to emit. Defaults to 'open-plugin' (.plugin/plugin.json). */
  manifestKind?: "open-plugin" | "claude-plugin" | "cursor-plugin";
}

export interface CopyOp {
  src: string;
  destRelative: string;
}

export interface MappedManifest {
  manifest: Record<string, unknown>;
  copyOps: CopyOp[];
  warnings: string[];
}

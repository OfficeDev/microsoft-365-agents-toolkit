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
}

export type AuthorizationType = "None" | "OAuthPluginVault" | "ApiKeyPluginVault";
export type DefaultAuthOption = "Auto" | AuthorizationType;

export interface ConvertInputs {
  path: string;
  output?: string;
  privacyUrl: string;
  termsUrl: string;
  websiteUrl?: string;
  appId?: string;
  defaultAuthType?: DefaultAuthOption;
  packageName?: string;
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

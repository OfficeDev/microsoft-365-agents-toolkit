// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Categorizes files processed during manifest resolution so that
 * telemetry and diagnostics can distinguish between different
 * artifact types in the app package.
 */
export enum ManifestType {
  TeamsManifest = "teams-manifest",
  PluginManifest = "plugin-manifest",
  DeclarativeCopilotManifest = "declarative-copilot-manifest",
  ApiSpec = "api-spec",
  EmbeddedKnowledgeFile = "embedded-knowledge-file",
}

/**
 * Result of resolving env-var placeholders in a manifest string.
 */
export interface EnvResolutionResult {
  /** The content with all placeholders resolved. */
  content: string;
  /** Placeholders that could not be resolved (empty on full success). */
  unresolved: EnvPlaceholder[];
}

/**
 * A single ${{VAR}} placeholder occurrence in manifest content.
 */
export interface EnvPlaceholder {
  /** The full token, e.g. "${{BOT_ID}}" */
  token: string;
  /** The variable name, e.g. "BOT_ID" */
  name: string;
}

/**
 * Common telemetry properties extracted from a Teams app manifest.
 * These are emitted with every manifest-related telemetry event so that
 * we can correlate app characteristics with operation outcomes.
 */
export interface ManifestTelemetryProperties {
  /** Teams app ID from the manifest. */
  id: string;
  /** App manifest version string (e.g. "1.0.0"). */
  version: string;
  /** Semicolon-joined list of capabilities (e.g. "staticTab;Bot;copilotGpt"). */
  capabilities: string;
  /** The $schema version (e.g. "1.17"). */
  manifestVersion: string;
  /** Whether it contains an API-based Message Extension. */
  isApiME: string;
  /** Whether it is a SharePoint Framework app. */
  isSPFx: string;
  /** Whether it is an API ME with Entra auth. */
  isApiMeAAD: string;
}

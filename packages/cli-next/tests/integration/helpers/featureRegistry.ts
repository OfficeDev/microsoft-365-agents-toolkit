/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 *
 * Typed wrapper around .dev/features.json.
 * Provides filter/resolve helpers for data-driven integration tests.
 * Adding a feature to features.json auto-creates new test cases.
 */

import * as fs from "node:fs";
import * as path from "node:path";

// ---------------------------------------------------------------------------
// Types mirroring features.json schema
// ---------------------------------------------------------------------------

export type FeatureLifecycle = "scaffold" | "provision" | "deploy" | "publish";
export type FeatureCategory =
  | "Bot"
  | "Tab"
  | "Custom Engine Agent"
  | "AI Agent"
  | "Declarative Agent"
  | "Connector"
  | "Messaging Extension"
  | "Office Add-in";

export interface FeatureEntry {
  id: string;
  name: string;
  description: string;
  category: FeatureCategory;
  templateName: string;
  languages: string[];
  lifecycles: FeatureLifecycle[];
  projectType: string;
  capabilities: string[];
  generatedFiles: string[];
  entryPoints: string[];
  adoSuiteId?: number;
  adoTestCaseCount?: number;
}

interface FeaturesFile {
  features: FeatureEntry[];
  trackedOnly: Array<{ id: string; name: string; templateName: string; reason: string }>;
}

// ---------------------------------------------------------------------------
// features.json templateName → actual folder name mapping
// ---------------------------------------------------------------------------

/**
 * Maps fx-core legacy templateName values (used in features.json) to
 * real folder names under templates/vsc/.
 * Only entries where the names differ are listed.
 */
const TEMPLATE_NAME_TO_FOLDER: Record<string, string> = {
  "copilot-gpt-basic": "declarative-agent-basic",
  "api-plugin-from-scratch": "declarative-agent-with-action-from-scratch",
  "api-plugin-from-scratch-oauth": "declarative-agent-with-action-from-scratch-oauth",
  "api-plugin-from-scratch-bearer": "declarative-agent-with-action-from-scratch-bearer",
  "non-sso-tab": "basic-tab",
  "default-message-extension": "message-extension-v2",
  "foundry-proxy-agent": "foundry-agent-to-m365",
};

/**
 * Resolve a features.json `templateName` to the actual folder name
 * under `templates/vsc/{lang}/`.
 */
export function resolveTemplateFolderName(templateName: string): string {
  return TEMPLATE_NAME_TO_FOLDER[templateName] ?? templateName;
}

// ---------------------------------------------------------------------------
// Loader
// ---------------------------------------------------------------------------

let _cached: FeaturesFile | undefined;

function loadFeaturesFile(): FeaturesFile {
  if (_cached) return _cached;
  const jsonPath = path.resolve(__dirname, "../../../../../.dev/features.json");
  const raw = fs.readFileSync(jsonPath, "utf-8");
  _cached = JSON.parse(raw) as FeaturesFile;
  return _cached;
}

// ---------------------------------------------------------------------------
// Query helpers
// ---------------------------------------------------------------------------

/** Get all testable feature entries. */
export function getFeatures(): FeatureEntry[] {
  return loadFeaturesFile().features;
}

/** Get features that include a specific lifecycle in their `lifecycles` array. */
export function getFeaturesForLifecycle(lifecycle: FeatureLifecycle): FeatureEntry[] {
  return getFeatures().filter((f) => f.lifecycles.includes(lifecycle));
}

/** Get features by category. */
export function getFeaturesByCategory(category: FeatureCategory): FeatureEntry[] {
  return getFeatures().filter((f) => f.category === category);
}

/** Get a single feature by ID. */
export function getFeatureById(id: string): FeatureEntry | undefined {
  return getFeatures().find((f) => f.id === id);
}

/**
 * Choose the best language variant for a feature's integration test.
 * Prefers: common → typescript → javascript → python → first available.
 */
export function preferredLanguage(feature: FeatureEntry): string {
  const pref = ["common", "typescript", "javascript", "python"];
  for (const lang of pref) {
    if (feature.languages.includes(lang)) return lang;
  }
  return feature.languages[0];
}

/**
 * Map a full language name to the folder key used under templates/vsc/.
 */
export function langToFolder(lang: string): string {
  switch (lang) {
    case "typescript":
      return "ts";
    case "javascript":
      return "js";
    case "csharp":
      return "csharp";
    default:
      return lang; // "python", "common"
  }
}

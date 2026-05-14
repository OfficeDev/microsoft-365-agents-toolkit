/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 *
 * Reads m365agents.yml.tpl from the template repository for a given
 * feature and language. Strips Mustache `{{...}}` placeholders so the
 * YAML is parseable by the lifecycle engine.
 */

import * as fs from "node:fs";
import * as path from "node:path";
import type { FeatureEntry } from "./featureRegistry";
import { resolveTemplateFolderName, preferredLanguage, langToFolder } from "./featureRegistry";

/**
 * Resolve the absolute path to a template's m365agents.yml.tpl file.
 * Returns `undefined` if the file doesn't exist on disk.
 */
export function resolveTemplateYamlPath(
  feature: FeatureEntry,
  language?: string
): string | undefined {
  const lang = language ?? preferredLanguage(feature);
  const folder = resolveTemplateFolderName(feature.templateName);
  const langDir = langToFolder(lang);

  // Navigate from this file → repo root → templates/vsc/{lang}/{folder}/
  const templatesRoot = path.resolve(__dirname, "../../../../../templates/vsc");
  const yamlPath = path.join(templatesRoot, langDir, folder, "m365agents.yml.tpl");

  return fs.existsSync(yamlPath) ? yamlPath : undefined;
}

/**
 * Read and preprocess the m365agents.yml.tpl file for a feature.
 *
 * Processing:
 * - Replaces `{{appName}}` and other Mustache vars with safe placeholder values
 * - Replaces `${{VAR}}` env references with literal strings so YAML parses cleanly
 *   (the executor will treat them as unresolved when resolving against the envMap)
 *
 * Returns `undefined` if the YAML file doesn't exist.
 */
export function loadTemplateYaml(feature: FeatureEntry, language?: string): string | undefined {
  const yamlPath = resolveTemplateYamlPath(feature, language);
  if (!yamlPath) return undefined;

  let content = fs.readFileSync(yamlPath, "utf-8");

  // Handle Mustache conditional blocks:
  // Keep content of {{#var}}...{{/var}} (positive branch)
  // Remove content of {{^var}}...{{/var}} (negative/else branch)
  // Process negative blocks first so we remove them before stripping positive tags
  content = content.replace(/\{\{\^([^}]+)\}\}[\s\S]*?\{\{\/\1\}\}/g, "");
  content = content.replace(/\{\{#([^}]+)\}\}/g, ""); // strip positive block open
  content = content.replace(/\{\{\/([^}]+)\}\}/g, ""); // strip block close

  // Replace simple Mustache vars with safe values
  content = content.replace(/\{\{appName\}\}/g, "test-app");
  content = content.replace(/\{\{([^}]+)\}\}/g, "placeholder-$1"); // eslint-disable-line no-secrets/no-secrets

  // Remove lines that became blank after Mustache stripping
  content = content
    .split("\n")
    .filter((line) => line.trim() !== "")
    .join("\n");

  return content;
}

/**
 * Scan ALL m365agents.yml.tpl files under templates/vsc/ and extract
 * unique `uses:` driver IDs per file. Used by driverCoverage tests.
 *
 * Returns a map of relative path → Set<driverID>.
 */
export function scanAllTemplateDrivers(): Map<string, Set<string>> {
  const templatesRoot = path.resolve(__dirname, "../../../../../templates/vsc");
  const result = new Map<string, Set<string>>();
  const usesPattern = /^\s*- uses:\s*(\S+)/;

  function walk(dir: string): void {
    const entries = fs.readdirSync(dir, { withFileTypes: true });
    for (const entry of entries) {
      const full = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        // Skip local, sandbox, playground variants
        if (["local", "sandbox", "playground"].includes(entry.name)) continue;
        walk(full);
      } else if (entry.name === "m365agents.yml.tpl") {
        const lines = fs.readFileSync(full, "utf-8").split("\n");
        const drivers = new Set<string>();
        for (const line of lines) {
          const match = usesPattern.exec(line);
          if (match) {
            // Strip trailing comments like "# Deploy given ARM templates..."
            drivers.add(match[1].replace(/#.*$/, "").trim());
          }
        }
        const relPath = path.relative(templatesRoot, full).replace(/\\/g, "/");
        result.set(relPath, drivers);
      }
    }
  }

  walk(templatesRoot);
  return result;
}

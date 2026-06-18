// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TemplateFileEntry } from "../model/dataModel";

/**
 * Narrow an opened package's flat `content/**` to the active language subtree
 * (ADR-0016 §5) — the single step between language-blind `open-template-package`
 * and the render phase.
 *
 * A language-agnostic package declares the reserved singleton `["common"]` and
 * ships one flat `content/` tree; it is returned verbatim (the two shipped
 * `["common"]` packages, `da/no-action` and `da/mcp-server`, are unaffected). A
 * language-partitioned package declares real languages and ships one subtree per
 * language under `content/{language}/`; this returns only the selected language's
 * entries with that prefix removed.
 *
 * The package readers are deliberately language-blind (they return every entry
 * flat so the registry existence check needs no language); this operation is the
 * one place that selects, after open and before render, so every consume path
 * (bundled-floor zip, on-disk authoring dir, T3 scenario harness) shares one rule.
 *
 * Spec: docs/03-specs/operations/scaffolding/select-language-content.md
 * (SLC-01..05)
 *
 * v4-owned (INV-7): imports no v3 symbol.
 */

/** The reserved single-member language axis for a language-agnostic template (ADR-0016 §5). */
export const COMMON_LANGUAGE = "common";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** Read the descriptor's declared `languages`, defaulting to the agnostic `["common"]`. */
function descriptorLanguages(descriptor: unknown): string[] {
  if (isRecord(descriptor) && Array.isArray(descriptor.languages)) {
    const languages = descriptor.languages.filter(
      (entry): entry is string => typeof entry === "string"
    );
    if (languages.length > 0) {
      return languages;
    }
  }
  return [COMMON_LANGUAGE];
}

/**
 * Select the content subtree the render phase should consume for `language`.
 *
 * @param descriptor the package's parsed `descriptor.json` (its `languages` axis)
 * @param content    the opened, deterministically ordered `content/**` entries
 * @param language   the selected language (Q0 answer; `"common"` when unprompted)
 */
export function selectLanguageContent(
  descriptor: unknown,
  content: TemplateFileEntry[],
  language: string
): TemplateFileEntry[] {
  // A `["common"]` package is never partitioned — keep its flat tree verbatim
  // (INV-1). The rule keys off the declared axis, not the floor value, so a
  // surface that passes a real project language for a `common` template still
  // gets the agnostic content (SLC-01/02).
  if (descriptorLanguages(descriptor).includes(COMMON_LANGUAGE)) {
    return content;
  }
  // A partitioned package keeps only the selected language's subtree, prefix
  // stripped, input order preserved (INV-2 / SLC-03..05).
  const prefix = `${language}/`;
  const selected: TemplateFileEntry[] = [];
  for (const entry of content) {
    if (entry.path.startsWith(prefix)) {
      selected.push({ path: entry.path.slice(prefix.length), data: entry.data });
    }
  }
  return selected;
}

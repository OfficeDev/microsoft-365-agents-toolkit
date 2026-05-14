// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { EnvPlaceholder } from "../types";

/**
 * Regex matching ${{VAR_NAME}} placeholders used in manifests and specs.
 * This is the toolkit-specific env-variable syntax that the manifest package
 * does not handle — it belongs in core-next.
 */
const PLACEHOLDER_RE = /\$\{\{([A-Za-z_][A-Za-z0-9_]*)\}\}/g;

/**
 * Scan a string (typically manifest JSON) for ${{VAR}} placeholders.
 *
 * @returns Unique placeholders found.
 */
export function findPlaceholders(text: string): EnvPlaceholder[] {
  const seen = new Set<string>();
  const results: EnvPlaceholder[] = [];
  let match: RegExpExecArray | null;

  while ((match = PLACEHOLDER_RE.exec(text)) !== null) {
    const name = match[1];
    if (!seen.has(name)) {
      seen.add(name);
      results.push({ token: match[0], name });
    }
  }

  return results;
}

/**
 * Resolve ${{VAR}} placeholders in text using the supplied env map.
 *
 * @param text   The manifest/spec content containing placeholders.
 * @param envMap A key→value map of environment variables.
 * @returns The resolved text and a list of any unresolved placeholders.
 */
export function resolvePlaceholders(
  text: string,
  envMap: ReadonlyMap<string, string> | Record<string, string>
): { resolved: string; unresolved: EnvPlaceholder[] } {
  const lookup: (key: string) => string | undefined =
    envMap instanceof Map ? (k) => envMap.get(k) : (k) => (envMap as Record<string, string>)[k];

  const unresolved: EnvPlaceholder[] = [];
  const resolved = text.replace(PLACEHOLDER_RE, (full, varName: string) => {
    const value = lookup(varName);
    if (value !== undefined) {
      return value;
    }
    unresolved.push({ token: full, name: varName });
    return full; // leave unresolved tokens intact
  });

  return { resolved, unresolved };
}

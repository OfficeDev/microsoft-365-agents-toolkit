// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { EnvPlaceholder } from "../declarativeAgent/types";

/**
 * Regex matching ${{VAR_NAME}} placeholders used in YAML config values.
 */
const PLACEHOLDER_RE = /\$\{\{([A-Za-z_][A-Za-z0-9_]*)\}\}/g;

/**
 * Resolve ${{VAR}} placeholders in driver step config values.
 *
 * Walks the config object recursively, replacing string placeholders
 * with values from the environment map.
 *
 * @param config  The driver `with` configuration from a YAML step.
 * @param envMap  Environment variable key→value mapping.
 * @returns A new config object with resolved values and a list of unresolved placeholders.
 */
export function resolveConfig(
  config: Record<string, unknown>,
  envMap: ReadonlyMap<string, string> | Record<string, string>
): { resolved: Record<string, unknown>; unresolved: EnvPlaceholder[] } {
  const lookup: (key: string) => string | undefined =
    envMap instanceof Map ? (k) => envMap.get(k) : (k) => (envMap as Record<string, string>)[k];

  const unresolved: EnvPlaceholder[] = [];

  function resolveValue(value: unknown): unknown {
    if (typeof value === "string") {
      return value.replace(PLACEHOLDER_RE, (full, varName: string) => {
        const resolved = lookup(varName);
        if (resolved !== undefined) {
          return resolved;
        }
        unresolved.push({ token: full, name: varName });
        return full;
      });
    }
    if (Array.isArray(value)) {
      return value.map(resolveValue);
    }
    if (value !== null && typeof value === "object") {
      const result: Record<string, unknown> = {};
      for (const [k, v] of Object.entries(value as Record<string, unknown>)) {
        result[k] = resolveValue(v);
      }
      return result;
    }
    return value;
  }

  const resolved = resolveValue(config) as Record<string, unknown>;
  return { resolved, unresolved };
}

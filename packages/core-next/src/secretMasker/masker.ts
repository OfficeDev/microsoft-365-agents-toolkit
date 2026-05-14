// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { matchesCredentialKeyword } from "./keywords";

const DEFAULT_MASK = "***";

// Matches "key=value" or "key:value" patterns where key contains a credential keyword.
// Value is captured as a quoted string or until a delimiter (whitespace, ;, &, comma).
const KEY_VALUE_REGEX = /(?<key>[\w.\-]+)(?:\s*[:=]\s*)(?<value>"[^"]*"|'[^']*'|[^\s;&,]+)/gi;

/**
 * Mask credential-like values in free-form text.
 * Scans for key=value patterns where the key matches a credential keyword.
 * Pure function — no side effects.
 *
 * @param text - Input text potentially containing secrets
 * @param replace - Replacement string (defaults to "***")
 * @returns Text with sensitive values masked
 */
export function maskSecret(text: string, replace: string = DEFAULT_MASK): string {
  if (!text) return text;
  return text.replace(KEY_VALUE_REGEX, (match, key, _value) => {
    if (matchesCredentialKeyword(key)) {
      return match.replace(_value, replace);
    }
    return match;
  });
}

/**
 * Mask values in a record whose keys match credential keywords.
 * Returns a new record — the original is not mutated.
 * Pure function — no side effects.
 *
 * @param record - Key-value pairs (e.g. env vars)
 * @param replace - Replacement string (defaults to "***")
 * @returns A new record with sensitive values replaced
 */
export function maskSecretValues(
  record: Record<string, string>,
  replace: string = DEFAULT_MASK
): Record<string, string> {
  const result: Record<string, string> = {};
  for (const [key, value] of Object.entries(record)) {
    result[key] = matchesCredentialKeyword(key) ? replace : value;
  }
  return result;
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Sanitisation utilities ported from the v3 CLI telemetry reporter.
 * These strip user-identifiable information before events reach App Insights.
 */

const fileRegex =
  /(file:\/\/)?([a-zA-Z]:(\\\\|\\|\/)|(\\\\|\\|\/))?([\w-._]{1,300}(\\\\|\\|\/)){1,100}[\w-._]{0,300}/g;
const nodeModulesRegex = /^[/\\]?(node_modules|node_modules\.asar)[/\\]/;

/**
 * Replace file-system paths in a string with a redaction marker,
 * preserving `node_modules/…` paths for stack-trace readability.
 */
export function anonymizeFilePaths(value: string | undefined): string {
  if (!value) return "";

  let result = "";
  let lastIndex = 0;

  // Reset stateful regex
  fileRegex.lastIndex = 0;

  let match: RegExpExecArray | null;
  while ((match = fileRegex.exec(value)) !== null) {
    if (!nodeModulesRegex.test(match[0])) {
      result += value.substring(lastIndex, match.index) + "<REDACTED: user-file-path>";
      lastIndex = fileRegex.lastIndex;
    }
  }

  if (lastIndex < value.length) {
    result += value.substring(lastIndex);
  }

  return result;
}

/**
 * Redact property values that may contain tokens, passwords, or email
 * addresses.
 */
export function sanitizeProperties(
  properties: Record<string, string> | undefined
): Record<string, string> | undefined {
  if (!properties) return properties;

  const emailRegex = /@[a-zA-Z0-9-.]+/;
  const cleaned: Record<string, string> = {};

  for (const [key, value] of Object.entries(properties)) {
    if (!value) continue;

    if (value.includes("token=")) {
      cleaned[key] = "<REDACTED: token>";
    } else if (value.includes("ssword=")) {
      cleaned[key] = "<REDACTED: password>";
    } else if (emailRegex.test(value)) {
      cleaned[key] = "<REDACTED: email>";
    } else {
      cleaned[key] = anonymizeFilePaths(value);
    }
  }

  return cleaned;
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { WhitelistFn } from "../../expression/evaluateExpression";

/** Closed v4 render-context function whitelist. See build-render-context spec. */

const NAMESPACE_FALLBACK = "mcpServer";
const NAMESPACE_MAX_LENGTH = 10;

/** Derive a stable MCP server namespace from a URL host. */
export function deriveMcpServerName(url: string): string {
  let host = "";
  try {
    host = new URL(url).host;
  } catch {
    host = "";
  }
  const alphanumeric = host.replace(/[^a-zA-Z0-9]/g, "").substring(0, NAMESPACE_MAX_LENGTH);
  return alphanumeric.length > 0 ? alphanumeric : NAMESPACE_FALLBACK;
}

/** Whitelist fn `mcpNamespace(url)` — the rendered `ai-plugin.json` namespace. */
export const mcpNamespace: WhitelistFn = (url: string): string => deriveMcpServerName(url);

/** Whitelist fn `mcpAuthRef(url)` for the deferred auth env-ref placeholder. */
export const mcpAuthRef: WhitelistFn = (url: string): string =>
  "${{MCP_DA_AUTH_ID_" + deriveMcpServerName(url).toUpperCase() + "}}";

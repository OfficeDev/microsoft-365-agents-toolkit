// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ExpressionRuntimePort, WhitelistFn } from "../expression/evaluateExpression";

/**
 * The closed, engine-owned whitelist of pure render-context functions
 * (ADR-0016 decision 3) and the real `ExpressionRuntimePort` that exposes them
 * plus the feature-flag reader behind `featureFlag('…')`.
 *
 * Spec: docs/03-specs/operations/scaffolding/build-render-context.md
 *       (the `{expr}` producers `mcpNamespace` / `mcpAuthRef` / `safeProjectNameLowerCase`)
 * Scenario: docs/03-specs/scenarios/da/create-mcp-server.md
 *           (SCN-CREATE-MCP-02 namespace, SCN-CREATE-MCP-05 reference_id),
 *           docs/03-specs/scenarios/da/create-api-plugin-from-scratch.md
 *           (SCN-CREATE-APIPLUGIN-01 the package.json name)
 *
 * v4-owned (INV-7): a self-contained reimplementation of the v3
 * `deriveMCPServerNameFromUrl` derivation — v4 imports no v3 symbol.
 */

const NAMESPACE_FALLBACK = "mcpServer";
const NAMESPACE_MAX_LENGTH = 10;

/**
 * Derive a stable MCP server namespace from its URL: the URL host with every
 * non-alphanumeric character stripped, truncated to ten characters, with a
 * fixed fallback when the URL has no usable host
 * (e.g. `https://api.github.com/…` → `apigithubc`).
 */
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

/**
 * Whitelist fn `mcpAuthRef(url)` — the `OAuthPluginVault` `reference_id`, an
 * env-ref placeholder `${{MCP_DA_AUTH_ID_<NS>}}` the render surface leaves
 * literal and provision resolves later.
 */
export const mcpAuthRef: WhitelistFn = (url: string): string =>
  "${{MCP_DA_AUTH_ID_" + deriveMcpServerName(url).toUpperCase() + "}}";

/**
 * Whitelist fn `safeProjectNameLowerCase(appName)` — the scaffolded backend's
 * `package.json` name: the app name with every non-alphanumeric character
 * stripped, lower-cased (e.g. `My Agent!` → `myagent`).
 *
 * v4-owned (INV-7): a self-contained reimplementation of the v3
 * `convertToAlphanumericOnly(appName).toLocaleLowerCase()` derivation. It uses
 * the locale-independent `toLowerCase` (not v3's `toLocaleLowerCase`) so the
 * whitelist stays pure and deterministic; the alphanumeric strip already leaves
 * only ASCII letters, for which the two agree.
 */
export const safeProjectNameLowerCase: WhitelistFn = (appName: string): string =>
  appName.replace(/[^0-9a-zA-Z]/g, "").toLowerCase();

/** The closed function whitelist; an author cannot extend it (ADR-0016 decision 3). */
const WHITELIST = new Map<string, WhitelistFn>([
  ["mcpNamespace", mcpNamespace],
  ["mcpAuthRef", mcpAuthRef],
  ["safeProjectNameLowerCase", safeProjectNameLowerCase],
]);

/** Default feature-flag reader: an env-backed truthy check (`"true"` / `"1"`). */
function envFlagReader(name: string): boolean {
  const value = process.env[name];
  return value === "true" || value === "1";
}

/**
 * Build the real `ExpressionRuntimePort`: the closed pure-function whitelist
 * plus a read-only feature-flag reader (env-backed by default; injected in
 * tests). Carries no `fs` / `http` / `clock` face (spec INV-4).
 */
export function createExpressionPort(
  flagReader: (name: string) => boolean = envFlagReader
): ExpressionRuntimePort {
  return {
    functions: (name: string): WhitelistFn | undefined => WHITELIST.get(name),
    flags: (name: string): boolean => flagReader(name),
  };
}

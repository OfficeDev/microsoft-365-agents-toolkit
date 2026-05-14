// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * E2E test configuration.
 *
 * Two modes:
 *   1. **CI mode** (`CI_ENABLED=true`) — reads credentials from environment
 *      variables (service principal or username/password).
 *   2. **Local mode** (default) — uses developer's cached credentials from
 *      `atk auth login`. No env vars required. Falls back to
 *      DefaultAzureCredential for Azure resource management.
 *
 * Local mode uses the Teams Cloud E2E Testing subscription by default:
 *   af46c703-f714-4f4c-af42-835a673c2b13 (TTL = 1 day)
 */

import {
  UsernamePasswordCredential,
  ClientSecretCredential,
  DefaultAzureCredential,
  type TokenCredential,
} from "@azure/identity";

/** Hardcoded client ID for E2E username/password auth (same as CLI). */
export const CLIENT_ID = "7ea7c24c-b1f6-4a20-9d11-9ae12e9e7ac0";

/** Default subscription for local developer runs (Teams Cloud – E2E Testing, TTL 1d). */
export const LOCAL_SUBSCRIPTION_ID = "af46c703-f714-4f4c-af42-835a673c2b13";

/** Whether running in CI (env-var-based auth) vs local (cached-credential auth). */
export function isCIMode(): boolean {
  return process.env.CI_ENABLED === "true";
}

/** Read optional env var. */
function optionalEnv(name: string): string | undefined {
  return process.env[name];
}

/** Read required env var — only enforced in CI mode. */
function requireEnv(name: string): string {
  const val = process.env[name];
  if (!val) {
    throw new Error(`E2E config: missing required environment variable ${name}`);
  }
  return val;
}

export interface E2EConfig {
  /** true when all credentials come from env vars (CI pipeline). */
  ciMode: boolean;
  azureTenantId: string;
  azureSubscriptionId: string;
  azureAccountName: string;
  azureAccountPassword: string;
  azureAccountObjectId: string;
  azureServicePrincipalId?: string;
  azureServicePrincipalSecret?: string;
  m365AccountName: string;
  m365AccountPassword: string;
  m365TenantId: string;
  githubRunId: string;
}

let _config: E2EConfig | undefined;

export function getConfig(): E2EConfig {
  if (!_config) {
    if (isCIMode()) {
      // CI mode — all credentials from env vars
      _config = {
        ciMode: true,
        azureTenantId: requireEnv("AZURE_TENANT_ID"),
        azureSubscriptionId: requireEnv("AZURE_SUBSCRIPTION_ID"),
        azureAccountName: requireEnv("AZURE_ACCOUNT_NAME"),
        azureAccountPassword: requireEnv("AZURE_ACCOUNT_PASSWORD"),
        azureAccountObjectId: optionalEnv("AZURE_ACCOUNT_OBJECT_ID") ?? "",
        azureServicePrincipalId: optionalEnv("AZURE_SERVICE_PRINCIPAL_ID"),
        azureServicePrincipalSecret: optionalEnv("AZURE_SERVICE_PRINCIPAL_SECRET"),
        m365AccountName: requireEnv("M365_ACCOUNT_NAME"),
        m365AccountPassword: requireEnv("M365_ACCOUNT_PASSWORD"),
        m365TenantId: requireEnv("M365_TENANT_ID"),
        githubRunId: optionalEnv("GITHUB_RUN_ID") ?? `local-${Date.now()}`,
      };
    } else {
      // Local mode — use developer's cached credentials from `atk auth login`
      _config = {
        ciMode: false,
        azureTenantId: optionalEnv("AZURE_TENANT_ID") ?? "",
        azureSubscriptionId: optionalEnv("AZURE_SUBSCRIPTION_ID") ?? LOCAL_SUBSCRIPTION_ID,
        azureAccountName: optionalEnv("AZURE_ACCOUNT_NAME") ?? "",
        azureAccountPassword: optionalEnv("AZURE_ACCOUNT_PASSWORD") ?? "",
        azureAccountObjectId: optionalEnv("AZURE_ACCOUNT_OBJECT_ID") ?? "",
        azureServicePrincipalId: optionalEnv("AZURE_SERVICE_PRINCIPAL_ID"),
        azureServicePrincipalSecret: optionalEnv("AZURE_SERVICE_PRINCIPAL_SECRET"),
        m365AccountName: optionalEnv("M365_ACCOUNT_NAME") ?? "",
        m365AccountPassword: optionalEnv("M365_ACCOUNT_PASSWORD") ?? "",
        m365TenantId: optionalEnv("M365_TENANT_ID") ?? "",
        githubRunId: `local-${Date.now()}`,
      };
    }
  }
  return _config;
}

/**
 * Create an Azure TokenCredential for resource management (RG create/delete).
 *
 * Priority:
 *   1. Service principal (CI)
 *   2. Username/password (CI)
 *   3. DefaultAzureCredential (local) — picks up `atk auth login` MSAL cache,
 *      Azure CLI, VS Code credential, etc.
 */
export function createAzureCredential(): TokenCredential {
  const cfg = getConfig();

  // CI: service principal
  if (cfg.azureServicePrincipalId && cfg.azureServicePrincipalSecret) {
    return new ClientSecretCredential(
      cfg.azureTenantId,
      cfg.azureServicePrincipalId,
      cfg.azureServicePrincipalSecret
    );
  }

  // CI: username/password
  if (cfg.ciMode && cfg.azureAccountName && cfg.azureAccountPassword) {
    return new UsernamePasswordCredential(
      cfg.azureTenantId,
      CLIENT_ID,
      cfg.azureAccountName,
      cfg.azureAccountPassword
    );
  }

  // Local: DefaultAzureCredential — includes MSAL shared cache, Azure CLI, etc.
  return new DefaultAzureCredential();
}

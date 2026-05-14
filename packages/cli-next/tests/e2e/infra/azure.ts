// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Azure resource management for E2E tests.
 * Uses resource tags instead of name-prefix matching for cleanup.
 *
 * Tags applied to every resource group:
 *   atk-test          = "true"
 *   test-run-id       = <GITHUB_RUN_ID or local UUID>
 *   test-template     = <template id, e.g. "bot/echo">
 *   created-at        = <ISO 8601 timestamp>
 *
 * Cleanup queries by tag, not by name pattern.
 */

import { ResourceManagementClient } from "@azure/arm-resources";
import { createAzureCredential, getConfig } from "./config";

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

let _client: ResourceManagementClient | undefined;

function getClient(): ResourceManagementClient {
  if (!_client) {
    const cfg = getConfig();
    _client = new ResourceManagementClient(createAzureCredential(), cfg.azureSubscriptionId);
  }
  return _client;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

export interface CreateResourceGroupOptions {
  name: string;
  location: string;
  templateId: string;
  runId: string;
  extraTags?: Record<string, string>;
}

/**
 * Create a resource group with E2E test tags.
 */
export async function createResourceGroup(opts: CreateResourceGroupOptions): Promise<boolean> {
  const client = getClient();
  try {
    await client.resourceGroups.createOrUpdate(opts.name, {
      location: opts.location,
      tags: {
        "atk-test": "true",
        "test-run-id": opts.runId,
        "test-template": opts.templateId,
        "created-at": new Date().toISOString(),
        ...opts.extraTags,
      },
    });
    console.log(`  [azure] Created resource group "${opts.name}"`);
    return true;
  } catch (e: unknown) {
    const msg = e instanceof Error ? e.message : String(e);
    console.error(`  [azure] Failed to create resource group "${opts.name}": ${msg}`);
    return false;
  }
}

/**
 * Delete a resource group with retries.
 */
export async function deleteResourceGroup(
  name: string,
  retries = 5,
  delayMs = 2000
): Promise<boolean> {
  const client = getClient();
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      const exists = await client.resourceGroups.checkExistence(name);
      if (!exists) {
        return true;
      }
      await client.resourceGroups.beginDeleteAndWait(name);
      console.log(`  [azure] Deleted resource group "${name}"`);
      return true;
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : String(e);
      if (attempt < retries) {
        console.warn(
          `  [azure] Delete "${name}" attempt ${attempt}/${retries} failed: ${msg}. Retrying...`
        );
        await delay(delayMs);
      } else {
        console.error(`  [azure] Failed to delete "${name}" after ${retries} attempts: ${msg}`);
      }
    }
  }
  return false;
}

/**
 * Find resource groups by tag filter.
 * Example: `listResourceGroupsByTag("atk-test", "true")` returns all E2E test RGs.
 */
export async function listResourceGroupsByTag(
  tagName: string,
  tagValue: string
): Promise<Array<{ name: string; tags: Record<string, string> }>> {
  const client = getClient();
  const results: Array<{ name: string; tags: Record<string, string> }> = [];
  for await (const rg of client.resourceGroups.list({
    filter: `tagName eq '${tagName}' and tagValue eq '${tagValue}'`,
  })) {
    if (rg.name) {
      results.push({ name: rg.name, tags: (rg.tags as Record<string, string>) ?? {} });
    }
  }
  return results;
}

/**
 * Delete all resource groups created by a specific test run.
 */
export async function deleteByRunId(runId: string): Promise<number> {
  const groups = await listResourceGroupsByTag("test-run-id", runId);
  let deleted = 0;
  await Promise.allSettled(
    groups.map(async (rg) => {
      const ok = await deleteResourceGroup(rg.name);
      if (ok) deleted++;
    })
  );
  if (groups.length > 0) {
    console.log(`  [azure] Cleaned ${deleted}/${groups.length} resource groups for run ${runId}`);
  }
  return deleted;
}

/**
 * Delete all E2E test resource groups older than `maxAgeHours`.
 */
export async function deleteStale(maxAgeHours: number): Promise<number> {
  const groups = await listResourceGroupsByTag("atk-test", "true");
  const cutoff = Date.now() - maxAgeHours * 60 * 60 * 1000;
  const stale = groups.filter((rg) => {
    const createdAt = rg.tags["created-at"];
    if (!createdAt) return true; // No timestamp → treat as stale
    return new Date(createdAt).getTime() < cutoff;
  });

  let deleted = 0;
  await Promise.allSettled(
    stale.map(async (rg) => {
      const ok = await deleteResourceGroup(rg.name);
      if (ok) deleted++;
    })
  );
  if (stale.length > 0) {
    console.log(
      `  [azure] Swept ${deleted}/${stale.length} stale resource groups (older than ${maxAgeHours}h)`
    );
  }
  return deleted;
}

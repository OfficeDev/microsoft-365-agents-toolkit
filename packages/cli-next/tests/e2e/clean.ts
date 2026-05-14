// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Stale resource cleanup script.
 * Run before/after E2E tests to sweep orphaned Azure resources.
 *
 * Usage:
 *   ts-node tests/e2e/clean.ts                    # delete resources older than 2h
 *   ts-node tests/e2e/clean.ts --run-id <id>      # delete resources from specific run
 */

import { deleteStale, deleteByRunId } from "./infra/azure";

async function main() {
  const args = process.argv.slice(2);
  const runIdIndex = args.indexOf("--run-id");

  if (runIdIndex !== -1 && args[runIdIndex + 1]) {
    const runId = args[runIdIndex + 1];
    console.log(`Cleaning resources for run ${runId}...`);
    await deleteByRunId(runId);
  } else {
    console.log("Sweeping stale E2E resources (older than 2 hours)...");
    await deleteStale(2);
  }
}

main().catch((e) => {
  console.error("Cleanup failed:", e);
  process.exitCode = 1;
});

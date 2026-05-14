// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Checkpoint-based retry: tracks which phases completed so Mocha's
 * `this.retries(1)` can resume from the failed phase instead of
 * re-running the entire test from scratch.
 *
 * On first run:       scaffold ✓ → provision ✓ → deploy ✗ (fails)
 * On retry (attempt 2): scaffold (skip) → provision (skip) → deploy (retry)
 *
 * State persists in a temp file so it survives Mocha's test re-instantiation.
 */

import * as fs from "fs";
import * as path from "path";
import * as os from "os";

export class TestCheckpoint {
  private state: Map<string, "done" | "failed"> = new Map();
  private stateFile: string;

  constructor(testId: string) {
    const dir = path.join(os.tmpdir(), "atk-e2e-checkpoints");
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
    this.stateFile = path.join(dir, `${testId.replace(/[/\\]/g, "_")}.json`);
    this.load();
  }

  /**
   * Run a phase. If already completed in a previous attempt, skip it.
   * If it fails, mark as failed and re-throw.
   */
  async runPhase<T>(name: string, fn: () => Promise<T>): Promise<T | undefined> {
    if (this.state.get(name) === "done") {
      console.log(`      ↪ [checkpoint] skipping "${name}" (completed in previous attempt)`);
      return undefined;
    }

    try {
      const result = await fn();
      this.state.set(name, "done");
      this.save();
      return result;
    } catch (err) {
      this.state.set(name, "failed");
      this.save();
      throw err;
    }
  }

  /** Check if a phase was already completed. */
  isDone(name: string): boolean {
    return this.state.get(name) === "done";
  }

  /** Reset all state (call in after() hook for clean runs). */
  reset(): void {
    this.state.clear();
    try {
      fs.unlinkSync(this.stateFile);
    } catch {
      // File may not exist
    }
  }

  private load(): void {
    try {
      if (fs.existsSync(this.stateFile)) {
        const data = JSON.parse(fs.readFileSync(this.stateFile, "utf-8"));
        this.state = new Map(Object.entries(data));
      }
    } catch {
      // Start fresh if corrupt
      this.state.clear();
    }
  }

  private save(): void {
    const obj: Record<string, string> = {};
    for (const [k, v] of this.state) {
      obj[k] = v;
    }
    fs.writeFileSync(this.stateFile, JSON.stringify(obj), "utf-8");
  }
}

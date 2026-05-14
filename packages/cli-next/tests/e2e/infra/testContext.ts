// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Creates an AtkContext wired for E2E testing:
 * - Real TokenProvider (from cli-next auth — same process, no subprocess)
 * - TestTracer as TelemetryReporter (captures all telemetry events)
 * - TestProgress as LifecycleProgress (captures step timing)
 * - Silent UI (auto-confirms, no prompts)
 * - Real file I/O and env management
 */

import * as crypto from "crypto";
import { createAtkContext, type AtkContext, LogProvider } from "@microsoft/teamsfx-core-next";
import { LogLevel } from "@microsoft/teamsfx-core-next";
import { createTokenProvider } from "../../../src/auth";
import { TestTracer, TestProgress } from "./tracer";
import { isCIMode } from "./config";
import { createCITokenProvider } from "./ciTokenProvider";

// ---------------------------------------------------------------------------
// Silent logger — writes to console but doesn't prompt
// ---------------------------------------------------------------------------

class SilentLogger implements LogProvider {
  log(logLevel: LogLevel, message: string): void {
    if (logLevel >= LogLevel.Info) {
      console.log(`  [${LogLevel[logLevel]}] ${message}`);
    }
  }
  verbose(_message: string): void {
    /* suppress */
  }
  debug(_message: string): void {
    /* suppress */
  }
  info(message: string | Array<{ content: string; color: any }>): void {
    const text = typeof message === "string" ? message : message.map((m) => m.content).join("");
    console.log(`  [info] ${text}`);
  }
  warning(message: string): void {
    console.warn(`  [warn] ${message}`);
  }
  error(message: string): void {
    console.error(`  [error] ${message}`);
  }
  async logInFile(_logLevel: LogLevel, _message: string): Promise<void> {
    /* no-op */
  }
  getLogFilePath(): string {
    return "/dev/null";
  }
}

// ---------------------------------------------------------------------------
// Silent UI — auto-confirms everything, no prompts
// ---------------------------------------------------------------------------

function createSilentUI() {
  return {
    selectOption: async () => ({ type: "success" as const, result: "" }),
    selectOptions: async () => ({ type: "success" as const, result: [] }),
    inputText: async () => ({ type: "success" as const, result: "" }),
    selectFile: async () => ({ type: "success" as const, result: "" }),
    selectFiles: async () => ({ type: "success" as const, result: [] }),
    selectFolder: async () => ({ type: "success" as const, result: "" }),
    openUrl: async () => ({ type: "success" as const, result: true }),
    showMessage: async () => undefined,
    createProgressBar: () => ({
      start: async () => {},
      next: async () => {},
      end: async () => {},
    }),
    confirm: async () => ({ type: "success" as const, result: true }),
  } as any;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

export interface TestContextResult {
  ctx: AtkContext;
  tracer: TestTracer;
  progress: TestProgress;
}

/**
 * Create an AtkContext for E2E testing with real auth and test instrumentation.
 *
 * @param projectPath - Path to the project being tested
 * @param correlationId - Optional correlation ID (defaults to random UUID)
 */
export function createTestContext(projectPath: string, correlationId?: string): TestContextResult {
  const tracer = new TestTracer();
  const progress = new TestProgress();

  const ctx = createAtkContext({
    auth: isCIMode() ? createCITokenProvider() : createTokenProvider(),
    logger: new SilentLogger(),
    telemetry: tracer,
    ui: createSilentUI(),
    projectPath,
    correlationId: correlationId ?? crypto.randomUUID(),
  });

  return { ctx, tracer, progress };
}

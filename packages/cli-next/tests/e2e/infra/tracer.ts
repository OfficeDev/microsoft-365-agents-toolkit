// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Structured test tracer — captures telemetry events and lifecycle progress
 * as OpenTelemetry-style trace spans. Writes JSONL for AI analysis.
 *
 * Also provides telemetry verification rules that detect instrumentation bugs.
 */

import * as fs from "fs";
import * as path from "path";
import type {
  TelemetryReporter,
  LifecycleProgress,
  LifecycleName,
} from "@microsoft/teamsfx-core-next";

// ---------------------------------------------------------------------------
// Trace types
// ---------------------------------------------------------------------------

export interface TraceSpan {
  name: string;
  timestamp: number;
  attributes: Record<string, string>;
  metrics: Record<string, number>;
  kind: "telemetry" | "telemetry-error" | "telemetry-exception" | "progress";
}

export interface StepTrace {
  index: number;
  label: string;
  startMs: number;
  durationMs?: number;
  status?: "success" | "failed";
}

export interface StepLog {
  test: string;
  step: string;
  status: "pass" | "fail" | "skip" | "retry";
  attempt: number;
  durationMs: number;
  command?: string;
  exitCode?: number;
  stdout?: string;
  stderr?: string;
  env?: Record<string, string>;
  assertions?: AssertionResult[];
  error?: { message: string; code?: string; stack?: string };
  telemetryIssues?: TelemetryIssue[];
}

export interface AssertionResult {
  name: string;
  passed: boolean;
  expected?: string;
  actual?: string;
  /** Failure severity. Defaults to "error" when omitted. */
  severity?: "error" | "warning";
  /**
   * Assertion tier:
   * - "shape"       — env-var presence, file existence, format regex
   * - "content"     — semantic parse / cross-field consistency (no network)
   * - "integration" — remote lookup (Graph, TDP, ARM, HTTP)
   * Defaults to "shape" when omitted.
   */
  tier?: "shape" | "content" | "integration";
}

// ---------------------------------------------------------------------------
// TestTracer — implements TelemetryReporter to capture events in-process
// ---------------------------------------------------------------------------

export class TestTracer implements TelemetryReporter {
  readonly spans: TraceSpan[] = [];

  sendTelemetryEvent(
    eventName: string,
    properties?: Record<string, string>,
    measurements?: Record<string, number>
  ): void {
    this.spans.push({
      name: eventName,
      timestamp: Date.now(),
      attributes: properties ?? {},
      metrics: measurements ?? {},
      kind: "telemetry",
    });
  }

  sendTelemetryErrorEvent(
    eventName: string,
    properties?: Record<string, string>,
    measurements?: Record<string, number>,
    _errorProps?: string[]
  ): void {
    this.spans.push({
      name: eventName,
      timestamp: Date.now(),
      attributes: properties ?? {},
      metrics: measurements ?? {},
      kind: "telemetry-error",
    });
  }

  sendTelemetryException(
    error: Error,
    properties?: Record<string, string>,
    measurements?: Record<string, number>
  ): void {
    this.spans.push({
      name: `exception:${error.name}`,
      timestamp: Date.now(),
      attributes: { ...properties, errorMessage: error.message },
      metrics: measurements ?? {},
      kind: "telemetry-exception",
    });
  }
}

// ---------------------------------------------------------------------------
// TestProgress — implements LifecycleProgress for step-level tracing
// ---------------------------------------------------------------------------

export class TestProgress implements LifecycleProgress {
  readonly steps: StepTrace[] = [];
  lifecycleName?: LifecycleName;
  totalSteps?: number;

  async onStart(lifecycle: LifecycleName, totalSteps: number): Promise<void> {
    this.lifecycleName = lifecycle;
    this.totalSteps = totalSteps;
  }

  async onStepStart(stepIndex: number, stepLabel: string): Promise<void> {
    this.steps.push({
      index: stepIndex,
      label: stepLabel,
      startMs: Date.now(),
    });
  }

  async onStepComplete(stepIndex: number, _stepLabel: string, durationMs: number): Promise<void> {
    const step = this.steps.find((s) => s.index === stepIndex);
    if (step) {
      step.durationMs = durationMs;
      step.status = "success";
    }
  }

  async onEnd(_success: boolean): Promise<void> {
    // Mark any incomplete steps as failed
    for (const step of this.steps) {
      if (!step.status) {
        step.status = "failed";
        step.durationMs = Date.now() - step.startMs;
      }
    }
  }
}

// ---------------------------------------------------------------------------
// StepLogger — writes structured JSONL log files
// ---------------------------------------------------------------------------

const SECRET_KEYS = [
  "SECRET",
  "PASSWORD",
  "TOKEN",
  "KEY",
  "CLIENT_SECRET",
  "CERTIFICATE",
  "CONNECTION_STRING",
];

function maskSecrets(env: Record<string, string>): Record<string, string> {
  const masked: Record<string, string> = {};
  for (const [k, v] of Object.entries(env)) {
    const isSecret = SECRET_KEYS.some((s) => k.toUpperCase().includes(s));
    masked[k] = isSecret ? "***" : v;
  }
  return masked;
}

function truncate(s: string | undefined, maxLen: number): string | undefined {
  if (!s) return undefined;
  return s.length > maxLen ? `...(truncated)...${s.slice(-maxLen)}` : s;
}

export class StepLogger {
  private entries: StepLog[] = [];
  private testName: string;
  private logDir: string;

  constructor(testName: string, logDir?: string) {
    this.testName = testName;
    this.logDir = logDir ?? path.join(__dirname, "..", "logs");
  }

  log(entry: StepLog): void {
    this.entries.push(entry);
    // Also write to console for live CI visibility
    const icon = entry.status === "pass" ? "✓" : entry.status === "fail" ? "✗" : "○";
    console.log(`    ${icon} [${entry.step}] ${entry.status} (${entry.durationMs}ms)`);
  }

  /**
   * Wrap a test step: times execution, catches errors, writes structured log.
   */
  async wrapStep(
    stepName: string,
    fn: () => Promise<{ assertions?: AssertionResult[]; env?: Record<string, string> } | void>,
    attempt = 1
  ): Promise<void> {
    const start = Date.now();
    try {
      const result = await fn();
      this.log({
        test: this.testName,
        step: stepName,
        status: "pass",
        attempt,
        durationMs: Date.now() - start,
        assertions: result?.assertions,
        env: result?.env ? maskSecrets(result.env) : undefined,
      });
    } catch (err: unknown) {
      const e = err instanceof Error ? err : new Error(String(err));
      this.log({
        test: this.testName,
        step: stepName,
        status: "fail",
        attempt,
        durationMs: Date.now() - start,
        error: {
          message: e.message,
          code: (e as any).code ?? (e as any).name,
          stack: truncate(e.stack, 1024),
        },
      });
      throw err; // Re-throw so Mocha sees the failure
    }
  }

  /**
   * Log telemetry verification results.
   */
  logTelemetryCheck(issues: TelemetryIssue[], totalSpans: number): void {
    this.log({
      test: this.testName,
      step: "telemetry-check",
      status: issues.length === 0 ? "pass" : "fail",
      attempt: 1,
      durationMs: 0,
      telemetryIssues: issues,
      assertions: [
        {
          name: `${totalSpans} telemetry spans verified`,
          passed: issues.length === 0,
          actual: `${issues.length} issues`,
        },
      ],
    });
  }

  /**
   * Flush all entries to a JSONL file on disk.
   */
  async flush(): Promise<string> {
    if (!fs.existsSync(this.logDir)) {
      fs.mkdirSync(this.logDir, { recursive: true });
    }
    const safeName = this.testName.replace(/[/\\]/g, "_");
    const filePath = path.join(this.logDir, `${safeName}-${Date.now()}.jsonl`);
    const lines = this.entries.map((e) => JSON.stringify(e)).join("\n") + "\n";
    await fs.promises.writeFile(filePath, lines, "utf-8");
    return filePath;
  }
}

// ---------------------------------------------------------------------------
// Telemetry verification rules
// ---------------------------------------------------------------------------

export interface TelemetryIssue {
  rule: string;
  message: string;
  spanName?: string;
}

export interface TelemetryRule {
  name: string;
  check(spans: TraceSpan[]): TelemetryIssue[];
}

/** Every *-start event must have a matching *-end event. */
const matchedStartEnd: TelemetryRule = {
  name: "matched-start-end",
  check(spans) {
    const starts = spans.filter((s) => s.name.endsWith("-start"));
    return starts
      .filter((s) => !spans.find((e) => e.name === s.name.replace("-start", "-end")))
      .map((s) => ({
        rule: "matched-start-end",
        message: `"${s.name}" has no matching end event`,
        spanName: s.name,
      }));
  },
};

/** End events must include success + duration properties. */
const requiredProperties: TelemetryRule = {
  name: "required-properties",
  check(spans) {
    return spans
      .filter((s) => s.name.endsWith("-end"))
      .filter((s) => !s.attributes.success || s.metrics.duration === undefined)
      .map((s) => ({
        rule: "required-properties",
        message: `"${s.name}" missing "success" attribute or "duration" metric`,
        spanName: s.name,
      }));
  },
};

/** A successful *-end event should not also be emitted as a telemetry-error. */
const noSpuriousErrors: TelemetryRule = {
  name: "no-spurious-errors",
  check(spans) {
    // Build a set of operation prefixes that ended successfully
    const successOps = new Set<string>();
    for (const s of spans) {
      if (s.name.endsWith("-end") && s.attributes.success === "true") {
        successOps.add(s.name.replace("-end", ""));
      }
    }
    // Flag error events whose operation ended successfully
    return spans
      .filter((s) => s.kind === "telemetry-error")
      .filter((s) => {
        const op = s.name.replace("-end", "").replace("-start", "");
        return successOps.has(op);
      })
      .map((s) => ({
        rule: "no-spurious-errors",
        message: `Error event "${s.name}" emitted for operation that ended successfully`,
        spanName: s.name,
      }));
  },
};

/** Duration must be non-negative and less than 30 minutes. */
const validDuration: TelemetryRule = {
  name: "valid-duration",
  check(spans) {
    return spans
      .filter((s) => s.metrics.duration !== undefined)
      .filter((s) => s.metrics.duration < 0 || s.metrics.duration > 30 * 60 * 1000)
      .map((s) => ({
        rule: "valid-duration",
        message: `"${s.name}" duration ${s.metrics.duration}ms is suspicious`,
        spanName: s.name,
      }));
  },
};

/** All events in a single operation should share the same correlationId. */
const consistentCorrelation: TelemetryRule = {
  name: "consistent-correlation",
  check(spans) {
    const ids = new Set(spans.map((s) => s.attributes.correlationId).filter(Boolean));
    if (ids.size > 1) {
      return [
        {
          rule: "consistent-correlation",
          message: `Multiple correlationIds found: [${[...ids].join(", ")}]`,
        },
      ];
    }
    return [];
  },
};

/** Error events must include errorCode. */
const errorHasCode: TelemetryRule = {
  name: "error-has-code",
  check(spans) {
    return spans
      .filter((s) => s.kind === "telemetry-error" || s.attributes.success === "false")
      .filter((s) => !s.attributes.errorCode)
      .map((s) => ({
        rule: "error-has-code",
        message: `"${s.name}" missing errorCode attribute`,
        spanName: s.name,
      }));
  },
};

export const BUILTIN_TELEMETRY_RULES: TelemetryRule[] = [
  matchedStartEnd,
  requiredProperties,
  noSpuriousErrors,
  validDuration,
  consistentCorrelation,
  errorHasCode,
];

/**
 * Run all telemetry rules against captured spans.
 * Returns an empty array if all rules pass.
 */
export function verifyTelemetry(
  spans: TraceSpan[],
  rules: TelemetryRule[] = BUILTIN_TELEMETRY_RULES
): TelemetryIssue[] {
  const issues: TelemetryIssue[] = [];
  for (const rule of rules) {
    issues.push(...rule.check(spans));
  }
  return issues;
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { LogProvider, TelemetryReporter, TokenProvider, CryptoProvider } from "../api/utils";
import { UserInteraction } from "../api/qm/ui";

/**
 * AtkContext is the dependency injection container that replaces the TOOLS singleton.
 * Every operation receives this as its first argument — no global state.
 *
 * Design principles:
 * - All external dependencies are injected here
 * - Each entry point (CLI, VS Code, Server) creates its own AtkContext
 * - Enables parallel testing by avoiding shared global state
 */
export interface AtkContext {
  /** Token acquisition for Azure and M365 */
  auth: TokenProvider;

  /** Structured logging */
  logger: LogProvider;

  /** Telemetry event tracking */
  telemetry: TelemetryReporter;

  /** User interaction (prompts, confirmations, progress) — platform-specific */
  ui: UserInteraction;

  /** Encrypt/decrypt secrets in env files */
  crypto?: CryptoProvider;

  /** Project root path (undefined when no project is open) */
  projectPath?: string;

  /** Correlation ID for tracing a single user action across components */
  correlationId: string;
}

/**
 * Create an AtkContext with required fields and sensible defaults.
 */
export function createAtkContext(params: {
  auth: TokenProvider;
  logger: LogProvider;
  telemetry: TelemetryReporter;
  ui: UserInteraction;
  crypto?: CryptoProvider;
  projectPath?: string;
  correlationId?: string;
}): AtkContext {
  return {
    auth: params.auth,
    logger: params.logger,
    telemetry: params.telemetry,
    ui: params.ui,
    crypto: params.crypto,
    projectPath: params.projectPath,
    correlationId: params.correlationId ?? generateCorrelationId(),
  };
}

function generateCorrelationId(): string {
  // Simple UUID v4 generation without external dependencies
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    const v = c === "x" ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });
}

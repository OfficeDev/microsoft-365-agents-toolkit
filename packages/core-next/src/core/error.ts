// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Simplified error model for v4.
 * Plain object with discriminated `kind` field — no class hierarchy.
 */
export interface AtkError {
  /** Stable error code for telemetry (e.g. "ManifestValidationError") */
  code: string;
  /** Human-readable error message */
  message: string;
  /** "user" = recoverable; "system" = infrastructure failure */
  kind: "user" | "system";
  /** Source component that raised the error */
  source?: string;
  /** Optional help link or mitigation guidance */
  help?: string;
  /** Optional inner error for root cause chain */
  inner?: Error;
  /** Optional display message for UI (may differ from technical message) */
  displayMessage?: string;
}

/**
 * Create a user error (recoverable, user-fixable).
 */
export function userError(
  code: string,
  message: string,
  options?: { source?: string; help?: string; inner?: Error; displayMessage?: string }
): AtkError {
  return { code, message, kind: "user", ...options };
}

/**
 * Create a system error (infrastructure/service failure).
 */
export function systemError(
  code: string,
  message: string,
  options?: { source?: string; help?: string; inner?: Error; displayMessage?: string }
): AtkError {
  return { code, message, kind: "system", ...options };
}

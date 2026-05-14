// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result, ok, err } from "neverthrow";
import { ZodSchema, ZodError } from "zod";
import { AtkContext } from "../core/context";
import { AtkError, userError } from "../core/error";
import { DriverConfig, DriverDescriptor, DriverOutput } from "./types";

/**
 * Options for creating a driver via the factory.
 */
export interface CreateDriverOptions<TConfig = DriverConfig> {
  /** Driver ID matching YAML `uses:` (e.g. "file/createOrUpdateEnvironmentFile") */
  id: string;
  /** Human-readable name for logging/UI */
  name: string;
  /** Zod schema for config validation (executed before `execute`) */
  inputSchema: ZodSchema<TConfig>;
  /** Driver execution logic — receives validated config */
  execute: (ctx: AtkContext, config: TConfig) => Promise<Result<DriverOutput, AtkError>>;
  /** Optional rollback logic */
  rollback?: (ctx: AtkContext, config: TConfig) => Promise<Result<void, AtkError>>;
}

/**
 * Create a `DriverDescriptor` with built-in Zod validation and error normalization.
 *
 * The returned descriptor:
 * 1. Validates `config` against the Zod schema before calling `execute`.
 * 2. Wraps unexpected exceptions into `AtkError` system errors.
 * 3. Sends start/end telemetry around execution.
 */
export function createDriver<TConfig = DriverConfig>(
  options: CreateDriverOptions<TConfig>
): DriverDescriptor {
  const { id, name, inputSchema, execute, rollback } = options;

  const executeFn = async (
    ctx: AtkContext,
    config: DriverConfig
  ): Promise<Result<DriverOutput, AtkError>> => {
    // --- Validate config ---
    const parsed = inputSchema.safeParse(config);
    if (!parsed.success) {
      return err(formatZodError(id, parsed.error));
    }

    // --- Execute with error boundary ---
    try {
      ctx.telemetry.sendTelemetryEvent("driver-start", { driver: id });
      const startMs = Date.now();
      const result = await execute(ctx, parsed.data);
      const durationMs = Date.now() - startMs;
      ctx.telemetry.sendTelemetryEvent(
        "driver-end",
        { driver: id, success: result.isOk() ? "true" : "false" },
        { duration: durationMs }
      );
      return result;
    } catch (error) {
      ctx.telemetry.sendTelemetryEvent(
        "driver-end",
        { driver: id, success: "false" },
        { duration: 0 }
      );
      return err(wrapUnexpectedError(id, error));
    }
  };

  const rollbackFn = rollback
    ? async (ctx: AtkContext, config: DriverConfig): Promise<Result<void, AtkError>> => {
        const parsed = inputSchema.safeParse(config);
        if (!parsed.success) {
          return err(formatZodError(id, parsed.error));
        }
        try {
          return await rollback(ctx, parsed.data);
        } catch (error) {
          return err(wrapUnexpectedError(id, error));
        }
      }
    : undefined;

  const validateFn = (config: DriverConfig): Result<void, AtkError> => {
    const parsed = inputSchema.safeParse(config);
    if (!parsed.success) {
      return err(formatZodError(id, parsed.error));
    }
    return ok(undefined);
  };

  return { id, name, executeFn, rollbackFn, validateFn };
}

/**
 * Transform a ZodError into a human-readable AtkError.
 */
function formatZodError(driverId: string, zodError: ZodError): AtkError {
  const issues = zodError.issues
    .map((issue) => `  - ${issue.path.join(".")}: ${issue.message}`)
    .join("\n");
  return userError(
    "InvalidDriverInput",
    `Driver "${driverId}" received invalid configuration:\n${issues}`,
    { source: driverId }
  );
}

/**
 * Wrap an unexpected exception into a system AtkError.
 * If the thrown value is already an AtkError (plain object with code/message/kind),
 * return it directly instead of wrapping into "[object Object]".
 */
function wrapUnexpectedError(driverId: string, error: unknown): AtkError {
  if (
    typeof error === "object" &&
    error !== null &&
    typeof (error as AtkError).code === "string" &&
    typeof (error as AtkError).message === "string" &&
    ((error as AtkError).kind === "user" || (error as AtkError).kind === "system")
  ) {
    return error as AtkError;
  }
  const message = error instanceof Error ? error.message : String(error);
  return {
    code: "DriverExecutionError",
    message: `Driver "${driverId}" failed unexpectedly: ${message}`,
    kind: "system",
    source: driverId,
    inner: error instanceof Error ? error : undefined,
  };
}

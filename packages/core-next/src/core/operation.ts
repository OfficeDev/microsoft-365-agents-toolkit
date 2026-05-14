// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result, err } from "neverthrow";
import { ZodSchema } from "zod";
import { AtkContext } from "./context";
import { AtkError, systemError } from "./error";

/**
 * An Operation is the fundamental unit of work in v4.
 * It replaces the @hooks middleware pattern with an explicit flat pipeline:
 *   Input → Validate (Zod) → Execute → Report
 *
 * Each operation is a standalone object — no class hierarchy, no hidden middleware.
 */
export interface Operation<TInput, TOutput> {
  /** Unique operation name, used for telemetry and logging */
  name: string;

  /** Zod schema that validates and transforms input before execution */
  inputSchema: ZodSchema<TInput>;

  /** The actual work. Receives validated input and context. */
  execute(ctx: AtkContext, input: TInput): Promise<Result<TOutput, AtkError>>;
}

/**
 * Run an operation with automatic validation, telemetry, and error wrapping.
 *
 * This is the thin wrapper replacing the old @hooks middleware chain.
 * The pipeline is explicit and traceable:
 * 1. Validate input against Zod schema
 * 2. Send telemetry start event
 * 3. Execute the operation
 * 4. Send telemetry end event (success or failure)
 * 5. Return result
 */
export async function runOperation<TInput, TOutput>(
  op: Operation<TInput, TOutput>,
  ctx: AtkContext,
  rawInput: unknown
): Promise<Result<TOutput, AtkError>> {
  // Step 1: Validate input
  const parseResult = op.inputSchema.safeParse(rawInput);
  if (!parseResult.success) {
    const zodError = parseResult.error;
    const message = zodError.issues
      .map((issue) => `${issue.path.join(".")}: ${issue.message}`)
      .join("; ");
    return err({
      code: "InputValidationError",
      message: `Invalid input for operation "${op.name}": ${message}`,
      kind: "user",
      source: op.name,
    });
  }

  const validInput = parseResult.data;

  // Step 2: Telemetry start
  const startTime = Date.now();
  ctx.telemetry.sendTelemetryEvent(`${op.name}-start`, {
    correlationId: ctx.correlationId,
  });
  ctx.logger.debug(`[${op.name}] Starting operation`);

  // Step 3: Execute
  let result: Result<TOutput, AtkError>;
  try {
    result = await op.execute(ctx, validInput);
  } catch (error) {
    result = err(
      systemError("UnhandledException", `Unhandled error in operation "${op.name}"`, {
        source: op.name,
        inner: error instanceof Error ? error : new Error(String(error)),
      })
    );
  }

  // Step 4: Telemetry end
  const duration = Date.now() - startTime;
  if (result.isOk()) {
    ctx.telemetry.sendTelemetryEvent(
      `${op.name}-end`,
      { correlationId: ctx.correlationId, success: "true" },
      { duration }
    );
    ctx.logger.debug(`[${op.name}] Completed successfully (${duration}ms)`);
  } else {
    ctx.telemetry.sendTelemetryErrorEvent(
      `${op.name}-end`,
      {
        correlationId: ctx.correlationId,
        success: "false",
        errorCode: result.error.code,
        errorKind: result.error.kind,
      },
      { duration }
    );
    ctx.logger.error(`[${op.name}] Failed: ${result.error.message} (${duration}ms)`);
  }

  return result;
}

/**
 * Helper to define an operation with type inference.
 */
export function defineOperation<TInput, TOutput>(
  name: string,
  inputSchema: ZodSchema<TInput>,
  execute: (ctx: AtkContext, input: TInput) => Promise<Result<TOutput, AtkError>>
): Operation<TInput, TOutput> {
  return { name, inputSchema, execute };
}

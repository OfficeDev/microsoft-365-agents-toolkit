// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result, err } from "neverthrow";
import { AtkContext } from "../core/context";
import { AtkError, systemError } from "../core/error";
import { sendStartEvent, sendSuccessEvent, sendErrorEvent } from "./helpers";

/**
 * Wrap an async function with start/end telemetry and timing.
 *
 * This is a standalone utility complementing `runOperation()` (which already
 * emits telemetry). Use `instrumentOperation` for ad-hoc work that doesn't
 * go through the Operation pipeline — e.g. subroutines, driver calls, etc.
 *
 * @param eventName - Telemetry event name
 * @param ctx - Current AtkContext
 * @param fn - Async work returning Result
 * @param properties - Optional extra properties sent with both start and end events
 */
export async function instrumentOperation<T>(
  eventName: string,
  ctx: AtkContext,
  fn: () => Promise<Result<T, AtkError>>,
  properties?: Record<string, string>
): Promise<Result<T, AtkError>> {
  sendStartEvent(ctx, eventName, properties);
  const start = Date.now();

  let result: Result<T, AtkError>;
  try {
    result = await fn();
  } catch (error) {
    const atkError = systemError(
      "UnhandledException",
      `Unhandled error in instrumented operation "${eventName}"`,
      { source: eventName, inner: error instanceof Error ? error : new Error(String(error)) }
    );
    const duration = Date.now() - start;
    sendErrorEvent(ctx, eventName, atkError, duration, properties);
    return err(atkError);
  }

  const duration = Date.now() - start;
  if (result.isOk()) {
    sendSuccessEvent(ctx, eventName, duration, properties);
  } else {
    sendErrorEvent(ctx, eventName, result.error, duration, properties);
  }
  return result;
}

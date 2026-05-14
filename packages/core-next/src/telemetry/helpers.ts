// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AtkContext } from "../core/context";
import { AtkError } from "../core/error";
import { TelemetryProperty, TelemetrySuccess } from "./types";
import { extractErrorProperties } from "./errorProperties";

/**
 * Send a start telemetry event for the given operation.
 * Emits "{eventName}-start" with correlation ID and optional extra properties.
 */
export function sendStartEvent(
  ctx: AtkContext,
  eventName: string,
  properties?: Record<string, string>
): void {
  ctx.telemetry.sendTelemetryEvent(`${eventName}-start`, {
    [TelemetryProperty.CorrelationId]: ctx.correlationId,
    [TelemetryProperty.OperationName]: eventName,
    ...properties,
  });
}

/**
 * Send a success telemetry event for the given operation.
 * Emits "{eventName}-end" with success=yes and duration.
 */
export function sendSuccessEvent(
  ctx: AtkContext,
  eventName: string,
  duration: number,
  properties?: Record<string, string>
): void {
  ctx.telemetry.sendTelemetryEvent(
    `${eventName}-end`,
    {
      [TelemetryProperty.CorrelationId]: ctx.correlationId,
      [TelemetryProperty.OperationName]: eventName,
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
      ...properties,
    },
    { [TelemetryProperty.Duration]: duration }
  );
}

/**
 * Send an error telemetry event for the given operation.
 * Emits "{eventName}-end" with success=no, error properties, and duration.
 */
export function sendErrorEvent(
  ctx: AtkContext,
  eventName: string,
  error: AtkError,
  duration: number,
  properties?: Record<string, string>
): void {
  ctx.telemetry.sendTelemetryErrorEvent(
    `${eventName}-end`,
    {
      [TelemetryProperty.CorrelationId]: ctx.correlationId,
      [TelemetryProperty.OperationName]: eventName,
      [TelemetryProperty.Success]: TelemetrySuccess.No,
      ...extractErrorProperties(error),
      ...properties,
    },
    { [TelemetryProperty.Duration]: duration }
  );
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AtkError } from "../core/error";
import { TelemetryProperty } from "./types";

/**
 * Extract telemetry-safe properties from an AtkError.
 * Pure function — no global state.
 */
export function extractErrorProperties(error: AtkError): Record<string, string> {
  const props: Record<string, string> = {
    [TelemetryProperty.ErrorCode]: error.code,
    [TelemetryProperty.ErrorKind]: error.kind,
    [TelemetryProperty.ErrorMessage]: error.message,
  };
  if (error.source) {
    props[TelemetryProperty.Component] = error.source;
  }
  if (error.inner) {
    props["inner-error-message"] = error.inner.message;
  }
  return props;
}

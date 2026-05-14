// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as crypto from "crypto";
import { createAtkContext, type AtkContext } from "@microsoft/teamsfx-core-next";
import { logger } from "./logger";
import { cliTelemetry } from "./telemetry";
import { cliUI } from "./ui";
import { createTokenProvider } from "./auth";

/**
 * Adapt the CLI telemetry reporter to the core TelemetryReporter interface.
 */
const telemetryAdapter = {
  sendTelemetryEvent(
    eventName: string,
    properties?: Record<string, string>,
    measurements?: Record<string, number>
  ): void {
    cliTelemetry.sendEvent(eventName, properties, measurements);
  },
  sendTelemetryErrorEvent(
    eventName: string,
    properties?: Record<string, string>,
    measurements?: Record<string, number>
  ): void {
    cliTelemetry.sendErrorEvent(eventName, new Error(eventName), properties, measurements);
  },
  sendTelemetryException(
    error: Error,
    properties?: Record<string, string>,
    measurements?: Record<string, number>
  ): void {
    cliTelemetry.sendErrorEvent("exception", error, properties, measurements);
  },
};

/**
 * Create an AtkContext for CLI command execution.
 *
 * @param projectPath - Optional project path (defaults to cwd)
 */
export function createCliContext(projectPath?: string): AtkContext {
  return createAtkContext({
    auth: createTokenProvider(),
    logger,
    telemetry: telemetryAdapter,
    ui: cliUI,
    projectPath: projectPath ?? process.cwd(),
    correlationId: crypto.randomUUID(),
  });
}

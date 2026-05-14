// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as os from "os";
import { AppInsightsTransport } from "./appInsightsTransport";
import { sanitizeProperties } from "./sanitize";

export interface TelemetryProperties {
  [key: string]: string;
}

export interface TelemetryMeasurements {
  [key: string]: number;
}

/**
 * CLI telemetry reporter backed by Application Insights.
 *
 * Lazy-initialised: until {@link init} is called with a non-empty key,
 * every method is a no-op.  This preserves the existing behaviour for
 * local / dev builds that ship without an `aiKey`.
 */
export class CliTelemetryReporter {
  private sharedProperties: TelemetryProperties = {};
  private transport: AppInsightsTransport | undefined;
  private cliName = "m365agentstoolkit-cli";
  private debug = false;

  /**
   * Activate telemetry.  Call once at startup.
   * If `aiKey` is falsy, no transport is created and all calls remain no-ops.
   */
  init(aiKey: string, cliVersion: string): void {
    if (!aiKey) return;

    this.debug =
      process.env.TEAMSFX_TELEMETRY_TEST === "true" || process.env.TEAMSFX_TELEMETRY_TEST === "1";

    let machineId = "unknown";
    try {
      const { machineIdSync } = require("node-machine-id") as {
        machineIdSync: () => string;
      };
      machineId = machineIdSync();
    } catch {
      // node-machine-id may fail in CI or containers — fall back silently
    }

    const commonProperties: Record<string, string> = {
      "common.os": os.platform(),
      "common.platformversion": (os.release() || "").replace(
        /^(\d+)(\.\d+)?(\.\d+)?(.*)/,
        "$1$2$3"
      ),
      "common.cliversion": cliVersion,
      "common.machineid": machineId,
    };

    this.transport = new AppInsightsTransport();
    this.transport.init(aiKey, commonProperties);
  }

  addSharedProperty(key: string, value: string): void {
    this.sharedProperties[key] = value;
  }

  sendEvent(
    eventName: string,
    properties?: TelemetryProperties,
    measurements?: TelemetryMeasurements
  ): void {
    if (!this.transport) return;

    const merged = { ...this.sharedProperties, ...properties };
    const cleaned = sanitizeProperties(merged);

    const fullName = `${this.cliName}/${eventName}`;
    this.transport.trackEvent(fullName, cleaned, measurements);

    if (this.debug) {
      console.debug(`[telemetry] ${fullName}`, { properties: cleaned, measurements });
    }
  }

  sendErrorEvent(
    eventName: string,
    error: Error,
    properties?: TelemetryProperties,
    measurements?: TelemetryMeasurements
  ): void {
    if (!this.transport) return;

    const merged = { ...this.sharedProperties, ...properties };
    const cleaned = sanitizeProperties(merged);

    const fullName = `${this.cliName}/${eventName}`;
    this.transport.trackEvent(fullName, cleaned, measurements);
    this.transport.trackException(error, cleaned, measurements);

    if (this.debug) {
      console.debug(`[telemetry] ${fullName}`, {
        error: error.message,
        properties: cleaned,
        measurements,
      });
    }
  }

  flush(): Promise<void> {
    if (!this.transport) return Promise.resolve();
    return this.transport.flush();
  }
}

/**
 * Singleton telemetry instance for the CLI.
 */
export const cliTelemetry = new CliTelemetryReporter();

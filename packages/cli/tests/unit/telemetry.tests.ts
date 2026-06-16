// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserCancelError } from "@microsoft/teamsfx-core";
import { afterEach, describe, expect, it, vi } from "vitest";
import { CliTelemetryReporter } from "../../src/commonlib/telemetry";
import cliTelemetry from "../../src/telemetry/cliTelemetry";

const validInstrumentationKey = "00000000-0000-0000-0000-000000000000";

describe("CLI Telemetry", function () {
  const stderrWrite = process.stderr.write.bind(process.stderr);

  beforeEach(() => {
    vi.spyOn(process.stderr, "write").mockImplementation(((chunk: any, ...args: any[]) => {
      const text = typeof chunk === "string" ? chunk : chunk?.toString?.() ?? "";
      if (text.includes("ApplicationInsights:An invalid instrumentation key was provided.")) {
        return true;
      }
      return stderrWrite(chunk, ...args);
    }) as any);
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });
  describe("disable", () => {
    it("no reporter", () => {
      cliTelemetry.enable = false;
    });
    it("sendTelemetryEvent", () => {
      cliTelemetry.reporter = new CliTelemetryReporter(
        validInstrumentationKey,
        "real",
        "real",
        "real"
      );
      const spy = vi.spyOn(cliTelemetry.reporter.reporter, "sendTelemetryEvent");
      cliTelemetry.enable = false;
      cliTelemetry.sendTelemetryEvent("eventName");
      expect(spy).not.toHaveBeenCalled();
    });
    it("sendTelemetryErrorEvent", () => {
      cliTelemetry.reporter = new CliTelemetryReporter(
        validInstrumentationKey,
        "real",
        "real",
        "real"
      );
      const spy = vi.spyOn(cliTelemetry.reporter.reporter, "sendTelemetryErrorEvent");
      cliTelemetry.enable = false;
      cliTelemetry.sendTelemetryErrorEvent("eventName", new UserCancelError());
      expect(spy).not.toHaveBeenCalled();
    });
    it("sendTelemetryException", () => {
      cliTelemetry.reporter = new CliTelemetryReporter(
        validInstrumentationKey,
        "real",
        "real",
        "real"
      );
      const spy = vi.spyOn(cliTelemetry.reporter.reporter, "sendTelemetryException");
      cliTelemetry.enable = false;
      cliTelemetry.sendTelemetryException(new Error());
      expect(spy).not.toHaveBeenCalled();
    });
  });
});

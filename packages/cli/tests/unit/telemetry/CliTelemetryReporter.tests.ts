// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as sinon from "sinon";
import { afterEach, describe, expect, it } from "vitest";
import { logger } from "../../../src/commonlib/logger";
import { CliTelemetryReporter } from "../../../src/commonlib/telemetry";

const validInstrumentationKey = "00000000-0000-0000-0000-000000000000";

describe("CliTelemetryReporter", () => {
  const sandbox = sinon.createSandbox();
  const stderrWrite = process.stderr.write.bind(process.stderr);

  beforeEach(() => {
    sandbox.stub(process.stderr, "write").callsFake(((chunk: any, ...args: any[]) => {
      const text = typeof chunk === "string" ? chunk : chunk?.toString?.() ?? "";
      if (text.includes("ApplicationInsights:An invalid instrumentation key was provided.")) {
        return true;
      }
      return stderrWrite(chunk, ...args);
    }) as any);
  });

  afterEach(() => {
    sandbox.restore();
  });

  describe("sendTelemetryErrorEvent", async () => {
    it("happy path", async () => {
      const reporter = new CliTelemetryReporter(validInstrumentationKey, "real", "real", "real");
      const debugStub = sandbox.stub(logger, "debug");
      const sendStub = sandbox.stub(reporter.reporter, "sendTelemetryErrorEvent");
      reporter.sendTelemetryErrorEvent("test");
      expect(debugStub.called).toBe(true);
      expect(sendStub.called).toBe(true);
    });
  });

  describe("sendTelemetryException", async () => {
    it("happy path", async () => {
      const reporter = new CliTelemetryReporter(validInstrumentationKey, "real", "real", "real");
      const stub = sandbox.stub(reporter.reporter, "sendTelemetryException");
      reporter.sendTelemetryException(new Error("test"));
      expect(stub.called).toBe(true);
    });
  });
});

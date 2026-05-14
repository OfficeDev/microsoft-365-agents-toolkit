/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, afterEach, beforeEach } from "mocha";
import * as sinon from "sinon";
import { CliTelemetryReporter } from "../../src/telemetry";
import { AppInsightsTransport } from "../../src/telemetry/appInsightsTransport";

describe("CliTelemetryReporter", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  describe("before init()", () => {
    it("sendEvent should be a no-op", () => {
      const reporter = new CliTelemetryReporter();
      // Should not throw
      reporter.sendEvent("test-event", { key: "value" }, { count: 1 });
    });

    it("sendErrorEvent should be a no-op", () => {
      const reporter = new CliTelemetryReporter();
      reporter.sendErrorEvent("test-error", new Error("boom"), { key: "value" });
    });

    it("flush should resolve immediately", async () => {
      const reporter = new CliTelemetryReporter();
      await reporter.flush();
    });
  });

  describe("init() with empty key", () => {
    it("should remain a no-op", () => {
      const reporter = new CliTelemetryReporter();
      reporter.init("", "1.0.0");
      // Still no-op — should not throw
      reporter.sendEvent("test-event");
    });
  });

  describe("after init() with valid key", () => {
    let reporter: CliTelemetryReporter;
    let transportInitStub: sinon.SinonStub;
    let trackEventStub: sinon.SinonStub;
    let trackExceptionStub: sinon.SinonStub;
    let flushStub: sinon.SinonStub;

    beforeEach(() => {
      // Stub AppInsightsTransport prototype before creating reporter
      transportInitStub = sandbox.stub(AppInsightsTransport.prototype, "init");
      trackEventStub = sandbox.stub(AppInsightsTransport.prototype, "trackEvent");
      trackExceptionStub = sandbox.stub(AppInsightsTransport.prototype, "trackException");
      flushStub = sandbox.stub(AppInsightsTransport.prototype, "flush").resolves();

      reporter = new CliTelemetryReporter();
      reporter.init("test-key-12345", "4.0.0");
    });

    it("should initialise transport with key and common properties", () => {
      expect(transportInitStub.calledOnce).to.be.true;
      const [key, commonProps] = transportInitStub.firstCall.args;
      expect(key).to.equal("test-key-12345");
      expect(commonProps).to.have.property("common.os");
      expect(commonProps).to.have.property("common.cliversion", "4.0.0");
      expect(commonProps).to.have.property("common.machineid");
    });

    it("sendEvent should track sanitized event", () => {
      reporter.addSharedProperty("binName", "atk");
      reporter.sendEvent("provision-start", { command: "provision" }, { duration: 100 });

      expect(trackEventStub.calledOnce).to.be.true;
      const [name, props, measurements] = trackEventStub.firstCall.args;
      expect(name).to.equal("m365agentstoolkit-cli/provision-start");
      expect(props).to.have.property("binName", "atk");
      expect(props).to.have.property("command", "provision");
      expect(measurements).to.deep.equal({ duration: 100 });
    });

    it("sendEvent should sanitize sensitive properties", () => {
      reporter.sendEvent("test", { auth: "token=abc123" });

      const [, props] = trackEventStub.firstCall.args;
      expect(props.auth).to.equal("<REDACTED: token>");
    });

    it("sendErrorEvent should track event and exception", () => {
      const err = new Error("test error");
      reporter.sendErrorEvent("deploy-error", err, { command: "deploy" }, { duration: 500 });

      expect(trackEventStub.calledOnce).to.be.true;
      expect(trackExceptionStub.calledOnce).to.be.true;

      const [, , exceptionMeasurements] = trackExceptionStub.firstCall.args;
      expect(exceptionMeasurements).to.deep.equal({ duration: 500 });
    });

    it("flush should delegate to transport", async () => {
      await reporter.flush();
      expect(flushStub.calledOnce).to.be.true;
    });
  });
});

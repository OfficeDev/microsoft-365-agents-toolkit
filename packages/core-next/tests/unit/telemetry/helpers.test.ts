/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, afterEach } from "mocha";
import * as sinon from "sinon";
import { sendStartEvent, sendSuccessEvent, sendErrorEvent } from "../../../src/telemetry/helpers";
import { TelemetryProperty, TelemetrySuccess } from "../../../src/telemetry/types";
import { userError } from "../../../src/core/error";
import { createMockContext } from "../testHelper";

describe("Telemetry helpers", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  describe("sendStartEvent()", () => {
    it("should send '{name}-start' event with correlation ID", () => {
      const ctx = createMockContext();
      sendStartEvent(ctx, "my-op");
      expect(ctx.telemetry.sendTelemetryEvent.calledOnce).to.be.true;
      const [eventName, props] = ctx.telemetry.sendTelemetryEvent.getCall(0).args;
      expect(eventName).to.equal("my-op-start");
      expect(props).to.include({
        [TelemetryProperty.CorrelationId]: "test-correlation-id-1234",
        [TelemetryProperty.OperationName]: "my-op",
      });
    });

    it("should include extra properties", () => {
      const ctx = createMockContext();
      sendStartEvent(ctx, "my-op", { custom: "value" });
      const [, props] = ctx.telemetry.sendTelemetryEvent.getCall(0).args;
      expect(props).to.include({ custom: "value" });
    });
  });

  describe("sendSuccessEvent()", () => {
    it("should send '{name}-end' with success=yes and duration", () => {
      const ctx = createMockContext();
      sendSuccessEvent(ctx, "my-op", 150);
      const [eventName, props, measurements] = ctx.telemetry.sendTelemetryEvent.getCall(0).args;
      expect(eventName).to.equal("my-op-end");
      expect(props).to.include({
        [TelemetryProperty.Success]: TelemetrySuccess.Yes,
      });
      expect(measurements).to.deep.include({
        [TelemetryProperty.Duration]: 150,
      });
    });
  });

  describe("sendErrorEvent()", () => {
    it("should send error event with error properties", () => {
      const ctx = createMockContext();
      const error = userError("BadInput", "test failure", { source: "test" });
      sendErrorEvent(ctx, "my-op", error, 200);
      expect(ctx.telemetry.sendTelemetryErrorEvent.calledOnce).to.be.true;
      const [eventName, props, measurements] =
        ctx.telemetry.sendTelemetryErrorEvent.getCall(0).args;
      expect(eventName).to.equal("my-op-end");
      expect(props).to.include({
        [TelemetryProperty.Success]: TelemetrySuccess.No,
        [TelemetryProperty.ErrorCode]: "BadInput",
        [TelemetryProperty.ErrorKind]: "user",
      });
      expect(measurements).to.deep.include({
        [TelemetryProperty.Duration]: 200,
      });
    });
  });
});

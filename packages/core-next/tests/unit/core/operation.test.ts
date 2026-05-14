/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, afterEach } from "mocha";
import * as sinon from "sinon";
import { z } from "zod";
import { ok, err } from "neverthrow";
import { runOperation, defineOperation } from "../../../src/core/operation";
import { userError } from "../../../src/core/error";
import { createMockContext } from "../testHelper";

describe("Operation", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  describe("defineOperation()", () => {
    it("should create an operation with name, schema, and execute", () => {
      const op = defineOperation("test-op", z.object({ name: z.string() }), async (_ctx, input) =>
        ok(input.name)
      );
      expect(op.name).to.equal("test-op");
      expect(op.inputSchema).to.exist;
      expect(op.execute).to.be.a("function");
    });
  });

  describe("runOperation()", () => {
    const testOp = defineOperation(
      "test-op",
      z.object({ value: z.number() }),
      async (_ctx, input) => ok(input.value * 2)
    );

    it("should validate input and execute successfully", async () => {
      const ctx = createMockContext();
      const result = await runOperation(testOp, ctx, { value: 21 });
      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.equal(42);
    });

    it("should return validation error for invalid input", async () => {
      const ctx = createMockContext();
      const result = await runOperation(testOp, ctx, { value: "not-a-number" });
      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("InputValidationError");
    });

    it("should send telemetry start and success events", async () => {
      const ctx = createMockContext();
      await runOperation(testOp, ctx, { value: 1 });
      expect(ctx.telemetry.sendTelemetryEvent.calledWith("test-op-start")).to.be.true;
      const endCall = ctx.telemetry.sendTelemetryEvent
        .getCalls()
        .find((c) => c.args[0] === "test-op-end");
      expect(endCall).to.exist;
      expect(endCall!.args[1]).to.include({ success: "true" });
    });

    it("should send telemetry error event on failure", async () => {
      const failOp = defineOperation("fail-op", z.object({}), async () =>
        err(userError("TestFail", "failed"))
      );
      const ctx = createMockContext();
      await runOperation(failOp, ctx, {});
      expect(ctx.telemetry.sendTelemetryErrorEvent.called).to.be.true;
      const errorCall = ctx.telemetry.sendTelemetryErrorEvent.getCall(0);
      expect(errorCall.args[1]).to.include({ success: "false", errorCode: "TestFail" });
    });

    it("should catch unhandled exceptions and wrap as system error", async () => {
      const throwOp = defineOperation("throw-op", z.object({}), async () => {
        throw new Error("boom");
      });
      const ctx = createMockContext();
      const result = await runOperation(throwOp, ctx, {});
      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("UnhandledException");
      expect(result._unsafeUnwrapErr().kind).to.equal("system");
    });
  });
});

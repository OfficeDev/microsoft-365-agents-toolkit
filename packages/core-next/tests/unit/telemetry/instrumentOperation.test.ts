/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, afterEach } from "mocha";
import * as sinon from "sinon";
import { ok, err } from "neverthrow";
import { instrumentOperation } from "../../../src/telemetry/instrumentOperation";
import { userError } from "../../../src/core/error";
import { createMockContext } from "../testHelper";

describe("instrumentOperation()", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("should send start and success events for ok result", async () => {
    const ctx = createMockContext();
    const result = await instrumentOperation("test-op", ctx, async () => ok(42));
    expect(result.isOk()).to.be.true;
    expect(result._unsafeUnwrap()).to.equal(42);
    expect(ctx.telemetry.sendTelemetryEvent.calledTwice).to.be.true;
    expect(ctx.telemetry.sendTelemetryEvent.getCall(0).args[0]).to.equal("test-op-start");
    expect(ctx.telemetry.sendTelemetryEvent.getCall(1).args[0]).to.equal("test-op-end");
  });

  it("should send start and error events for err result", async () => {
    const ctx = createMockContext();
    const error = userError("Fail", "failed");
    const result = await instrumentOperation("test-op", ctx, async () => err(error));
    expect(result.isErr()).to.be.true;
    expect(ctx.telemetry.sendTelemetryEvent.calledOnce).to.be.true; // start only
    expect(ctx.telemetry.sendTelemetryErrorEvent.calledOnce).to.be.true; // error end
  });

  it("should catch thrown exceptions and return system error", async () => {
    const ctx = createMockContext();
    const result = await instrumentOperation("test-op", ctx, async () => {
      throw new Error("boom");
    });
    expect(result.isErr()).to.be.true;
    expect(result._unsafeUnwrapErr().code).to.equal("UnhandledException");
    expect(result._unsafeUnwrapErr().kind).to.equal("system");
  });

  it("should pass extra properties to telemetry events", async () => {
    const ctx = createMockContext();
    await instrumentOperation("test-op", ctx, async () => ok("done"), { custom: "prop" });
    const [, startProps] = ctx.telemetry.sendTelemetryEvent.getCall(0).args;
    expect(startProps).to.include({ custom: "prop" });
  });
});

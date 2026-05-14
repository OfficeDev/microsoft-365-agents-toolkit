/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it } from "mocha";
import * as sinon from "sinon";
import { ok, err } from "neverthrow";
import { z } from "zod";
import { createDriver } from "../../../src/drivers/createDriver";
import { userError } from "../../../src/core/error";
import { createMockContext } from "../testHelper";

describe("createDriver", () => {
  const schema = z.object({
    name: z.string(),
    count: z.number().optional(),
  });

  it("should create a valid DriverDescriptor", () => {
    const driver = createDriver({
      id: "test/driver",
      name: "Test Driver",
      inputSchema: schema,
      execute: async () => ok({ outputs: {} }),
    });

    expect(driver.id).to.equal("test/driver");
    expect(driver.name).to.equal("Test Driver");
    expect(driver.executeFn).to.be.a("function");
    expect(driver.validateFn).to.be.a("function");
    expect(driver.rollbackFn).to.be.undefined;
  });

  it("should validate config before execution", async () => {
    const executeFn = sinon.stub().resolves(ok({ outputs: {} }));
    const driver = createDriver({
      id: "test/driver",
      name: "Test Driver",
      inputSchema: schema,
      execute: executeFn,
    });

    const ctx = createMockContext();
    const result = await driver.executeFn(ctx, { name: "hello" });

    expect(result.isOk()).to.be.true;
    expect(executeFn.calledOnce).to.be.true;
    expect(executeFn.firstCall.args[1]).to.deep.equal({ name: "hello" });
  });

  it("should return validation error for invalid config", async () => {
    const executeFn = sinon.stub().resolves(ok({ outputs: {} }));
    const driver = createDriver({
      id: "test/driver",
      name: "Test Driver",
      inputSchema: schema,
      execute: executeFn,
    });

    const ctx = createMockContext();
    const result = await driver.executeFn(ctx, { count: "not-a-number" });

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("InvalidDriverInput");
      expect(result.error.kind).to.equal("user");
      expect(result.error.message).to.include("test/driver");
    }
    expect(executeFn.called).to.be.false;
  });

  it("should catch unexpected exceptions from execute", async () => {
    const driver = createDriver({
      id: "test/driver",
      name: "Test Driver",
      inputSchema: schema,
      execute: async () => {
        throw new Error("boom");
      },
    });

    const ctx = createMockContext();
    const result = await driver.executeFn(ctx, { name: "hello" });

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("DriverExecutionError");
      expect(result.error.kind).to.equal("system");
      expect(result.error.message).to.include("boom");
    }
  });

  it("should send telemetry events around execution", async () => {
    const driver = createDriver({
      id: "test/driver",
      name: "Test Driver",
      inputSchema: schema,
      execute: async () => ok({ outputs: { key: "value" } }),
    });

    const ctx = createMockContext();
    await driver.executeFn(ctx, { name: "hello" });

    const telemetry = ctx.telemetry;
    const startCall = telemetry.sendTelemetryEvent
      .getCalls()
      .find((c: sinon.SinonSpyCall) => c.args[0] === "driver-start");
    expect(startCall).to.exist;
    expect(startCall!.args[1]).to.deep.equal({ driver: "test/driver" });

    const endCall = telemetry.sendTelemetryEvent
      .getCalls()
      .find((c: sinon.SinonSpyCall) => c.args[0] === "driver-end");
    expect(endCall).to.exist;
    expect(endCall!.args[1]).to.deep.equal({ driver: "test/driver", success: "true" });
    expect(endCall!.args[2]).to.have.property("duration").that.is.a("number");
  });

  it("should pass through driver errors as-is", async () => {
    const driverError = userError("CustomError", "something went wrong");
    const driver = createDriver({
      id: "test/driver",
      name: "Test Driver",
      inputSchema: schema,
      execute: async () => err(driverError),
    });

    const ctx = createMockContext();
    const result = await driver.executeFn(ctx, { name: "hello" });

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error).to.deep.equal(driverError);
    }
  });

  it("validateFn should succeed for valid config", () => {
    const driver = createDriver({
      id: "test/driver",
      name: "Test Driver",
      inputSchema: schema,
      execute: async () => ok({ outputs: {} }),
    });

    const result = driver.validateFn!({ name: "hello", count: 5 });
    expect(result.isOk()).to.be.true;
  });

  it("validateFn should fail for invalid config", () => {
    const driver = createDriver({
      id: "test/driver",
      name: "Test Driver",
      inputSchema: schema,
      execute: async () => ok({ outputs: {} }),
    });

    const result = driver.validateFn!({ bad: true });
    expect(result.isErr()).to.be.true;
  });

  it("should create rollbackFn when rollback is provided", async () => {
    const rollbackStub = sinon.stub().resolves(ok(undefined));
    const driver = createDriver({
      id: "test/driver",
      name: "Test Driver",
      inputSchema: schema,
      execute: async () => ok({ outputs: {} }),
      rollback: rollbackStub,
    });

    expect(driver.rollbackFn).to.be.a("function");

    const ctx = createMockContext();
    const result = await driver.rollbackFn!(ctx, { name: "hello" });

    expect(result.isOk()).to.be.true;
    expect(rollbackStub.calledOnce).to.be.true;
  });

  it("rollbackFn should validate config before rollback", async () => {
    const rollbackStub = sinon.stub().resolves(ok(undefined));
    const driver = createDriver({
      id: "test/driver",
      name: "Test Driver",
      inputSchema: schema,
      execute: async () => ok({ outputs: {} }),
      rollback: rollbackStub,
    });

    const ctx = createMockContext();
    const result = await driver.rollbackFn!(ctx, {});

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("InvalidDriverInput");
    }
    expect(rollbackStub.called).to.be.false;
  });
});

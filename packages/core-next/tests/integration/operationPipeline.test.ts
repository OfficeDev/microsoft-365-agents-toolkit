/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, afterEach } from "mocha";
import * as sinon from "sinon";
import { z } from "zod";
import { ok, err } from "neverthrow";
import { runOperation, defineOperation, createAtkContext, userError } from "../../src";
import {
  instrumentOperation,
  correlationScope,
  getCurrentCorrelationId,
} from "../../src/telemetry";

/**
 * Integration test: exercises the full operation pipeline with telemetry helpers,
 * correlation scoping, and error handling working together end-to-end.
 * No external services needed — all deps are mocked via AtkContext.
 */
describe("Integration: Operation → Telemetry pipeline", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  function createTestContext() {
    return createAtkContext({
      auth: { m365TokenProvider: {} as any, azureAccountProvider: {} as any },
      logger: {
        log: sinon.stub(),
        verbose: sinon.stub(),
        debug: sinon.stub(),
        info: sinon.stub(),
        warning: sinon.stub(),
        error: sinon.stub(),
        logInFile: sinon.stub().resolves(),
        getLogFilePath: sinon.stub().returns("/tmp/test.log"),
      } as any,
      telemetry: {
        sendTelemetryEvent: sinon.stub(),
        sendTelemetryErrorEvent: sinon.stub(),
        sendTelemetryException: sinon.stub(),
      },
      ui: {
        selectOption: sinon.stub(),
        selectOptions: sinon.stub(),
        inputText: sinon.stub(),
        selectFile: sinon.stub(),
        selectFiles: sinon.stub(),
        selectFolder: sinon.stub(),
        openUrl: sinon.stub(),
        showMessage: sinon.stub(),
        createProgressBar: sinon.stub(),
        confirm: sinon.stub(),
      } as any,
      correlationId: "integration-test-id",
    });
  }

  it("should run a successful operation and emit correct telemetry sequence", async () => {
    const addOp = defineOperation(
      "add-numbers",
      z.object({ a: z.number(), b: z.number() }),
      async (_ctx, input) => ok(input.a + input.b)
    );

    const ctx = createTestContext();
    const result = await runOperation(addOp, ctx, { a: 3, b: 4 });

    expect(result.isOk()).to.be.true;
    expect(result._unsafeUnwrap()).to.equal(7);

    // Verify telemetry sequence: start → end(success)
    const tel = ctx.telemetry as sinon.SinonStubbedInstance<typeof ctx.telemetry>;
    expect(tel.sendTelemetryEvent.callCount).to.be.greaterThanOrEqual(2);
    expect(tel.sendTelemetryEvent.getCall(0).args[0]).to.equal("add-numbers-start");
    expect(tel.sendTelemetryEvent.getCall(1).args[0]).to.equal("add-numbers-end");
    expect(tel.sendTelemetryEvent.getCall(1).args[1]).to.include({ success: "true" });
  });

  it("should run a failing operation and emit error telemetry", async () => {
    const failOp = defineOperation("fail-op", z.object({}), async () =>
      err(userError("Kaboom", "something broke"))
    );

    const ctx = createTestContext();
    const result = await runOperation(failOp, ctx, {});

    expect(result.isErr()).to.be.true;
    expect(result._unsafeUnwrapErr().code).to.equal("Kaboom");

    const tel = ctx.telemetry as sinon.SinonStubbedInstance<typeof ctx.telemetry>;
    expect(tel.sendTelemetryErrorEvent.calledOnce).to.be.true;
    expect(tel.sendTelemetryErrorEvent.getCall(0).args[1]).to.include({
      errorCode: "Kaboom",
      errorKind: "user",
    });
  });

  it("should propagate correlation ID through nested operations", async () => {
    const ids: string[] = [];

    const outerOp = defineOperation("outer-op", z.object({}), async (ctx, _input) => {
      return correlationScope(ctx.correlationId, () => {
        ids.push(getCurrentCorrelationId());
        return Promise.resolve(ok("done"));
      });
    });

    const ctx = createTestContext();
    await runOperation(outerOp, ctx, {});

    expect(ids).to.deep.equal(["integration-test-id"]);
  });

  it("should validate input before executing and emit no success telemetry", async () => {
    const strictOp = defineOperation(
      "strict-op",
      z.object({ name: z.string().min(3) }),
      async (_ctx, input) => ok(input.name)
    );

    const ctx = createTestContext();
    const result = await runOperation(strictOp, ctx, { name: "ab" }); // too short

    expect(result.isErr()).to.be.true;
    expect(result._unsafeUnwrapErr().code).to.equal("InputValidationError");

    // No end telemetry emitted for validation failures (start is not sent either)
    const tel = ctx.telemetry as sinon.SinonStubbedInstance<typeof ctx.telemetry>;
    expect(tel.sendTelemetryErrorEvent.called).to.be.false;
  });

  it("instrumentOperation should compose with existing telemetry context", async () => {
    const ctx = createTestContext();

    const result = await instrumentOperation("sub-task", ctx, async () => {
      return ok({ data: "processed" });
    });

    expect(result.isOk()).to.be.true;

    const tel = ctx.telemetry as sinon.SinonStubbedInstance<typeof ctx.telemetry>;
    // instrumentOperation sends its own start + end
    const events = tel.sendTelemetryEvent.getCalls().map((c) => c.args[0]);
    expect(events).to.include("sub-task-start");
    expect(events).to.include("sub-task-end");
  });
});

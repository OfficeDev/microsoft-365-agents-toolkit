// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * @author Siglud <siglud@gmail.com>
 */
import { HookContext, NextFunction } from "@feathersjs/hooks";
import { FxError } from "@microsoft/teamsfx-api";
import { expect } from "chai";
import * as sinon from "sinon";
import { addSWADeployTelemetry } from "../../../src/component/driver/middleware/addSWADeployTelemetry";

describe("addSWADeployTelemetry", () => {
  let clock: sinon.SinonFakeTimers;
  let next: NextFunction;
  let ctx: HookContext;
  let telemetryReporter: {
    sendTelemetryEvent: sinon.SinonSpy;
    sendTelemetryErrorEvent: sinon.SinonSpy;
  };

  beforeEach(() => {
    clock = sinon.useFakeTimers();
    next = sinon.spy(() => {
      return Promise.resolve(12);
    }) as NextFunction;
    telemetryReporter = {
      sendTelemetryEvent: sinon.spy(),
      sendTelemetryErrorEvent: sinon.spy(),
    };
    ctx = {
      arguments: [
        { args: "test command" },
        { telemetryReporter, telemetryProperties: {} } as any,
        undefined,
        "",
        "deploy to Azure Static Web Apps",
      ],
      result: { isOk: sinon.fake.returns(true) },
    } as HookContext;
  });

  afterEach(() => {
    clock.restore();
    sinon.restore();
  });

  it("should not add telemetry for script", async () => {
    const middleware = addSWADeployTelemetry("testEvent");
    const res = await middleware(ctx, next);
    expect(res).to.equal(undefined);
  });

  it("should add telemetry for non-script", async () => {
    const middleware = addSWADeployTelemetry("testEvent");
    await middleware(ctx, next);
    clock.tick(1000); // Simulate time passing
    expect(telemetryReporter.sendTelemetryEvent.called).to.be.true;
    expect(telemetryReporter.sendTelemetryErrorEvent.called).to.be.false;
  });

  it("When name is not deploy to Azure Static Web Apps", async () => {
    const ctx = {
      arguments: [
        { args: "test command" },
        { telemetryReporter, telemetryProperties: {} } as any,
        undefined,
        "",
        "Anything else",
      ],
      result: { isOk: sinon.fake.returns(true) },
    } as HookContext;
    const middleware = addSWADeployTelemetry("testEvent");
    const res = await middleware(ctx, next);
    expect(res).to.equal(undefined);
    expect(telemetryReporter.sendTelemetryEvent.called).to.be.false;
    expect(telemetryReporter.sendTelemetryErrorEvent.called).to.be.false;
  });

  it("When return value is not ok", async () => {
    const err = { e: "error" } as unknown as FxError;
    const ctx = {
      arguments: [
        { args: "test command" },
        { telemetryReporter, telemetryProperties: {} } as any,
        undefined,
        "",
        "deploy to Azure Static Web Apps",
      ],
      result: { isOk: sinon.fake.returns(false), error: err },
    } as HookContext;
    const middleware = addSWADeployTelemetry("testEvent");
    await middleware(ctx, next);
    expect(telemetryReporter.sendTelemetryEvent.called).to.be.true;
    expect(telemetryReporter.sendTelemetryErrorEvent.called).to.be.true;
  });
});

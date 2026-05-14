/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, afterEach, beforeEach } from "mocha";
import * as sinon from "sinon";
import { Command } from "commander";
import { wrapHandler } from "../../src/handler";
import { cliTelemetry } from "../../src/telemetry";

describe("wrapHandler()", () => {
  const sandbox = sinon.createSandbox();
  let sendEventStub: sinon.SinonStub;
  let sendErrorEventStub: sinon.SinonStub;
  let flushStub: sinon.SinonStub;

  beforeEach(() => {
    sendEventStub = sandbox.stub(cliTelemetry, "sendEvent");
    sendErrorEventStub = sandbox.stub(cliTelemetry, "sendErrorEvent");
    flushStub = sandbox.stub(cliTelemetry, "flush").resolves();
    sandbox.stub(console, "error");
  });

  afterEach(() => {
    sandbox.restore();
    process.exitCode = undefined;
  });

  function makeMockCmd(opts: Record<string, unknown> = {}): Command {
    const cmd = new Command("test-cmd");
    for (const [k, v] of Object.entries(opts)) {
      cmd.setOptionValue(k, v);
    }
    return cmd;
  }

  it("should send start and end telemetry on success", async () => {
    const handler = sandbox.stub().resolves();
    const wrapped = wrapHandler("mycommand", handler);

    await wrapped(makeMockCmd());

    expect(sendEventStub.calledTwice).to.be.true;
    expect(sendEventStub.getCall(0).args[0]).to.equal("mycommand-start");
    expect(sendEventStub.getCall(1).args[0]).to.equal("mycommand-end");
    expect(flushStub.calledOnce).to.be.true;
  });

  it("should send error telemetry and set exitCode on failure", async () => {
    const err = new Error("boom");
    const handler = sandbox.stub().rejects(err);
    const wrapped = wrapHandler("failcmd", handler);

    await wrapped(makeMockCmd());

    expect(sendEventStub.calledOnce).to.be.true;
    expect(sendEventStub.getCall(0).args[0]).to.equal("failcmd-start");
    expect(sendErrorEventStub.calledOnce).to.be.true;
    expect(sendErrorEventStub.getCall(0).args[0]).to.equal("failcmd-error");
    expect(process.exitCode).to.equal(1);
  });

  it("should call flush even on error", async () => {
    const handler = sandbox.stub().rejects(new Error("oops"));
    const wrapped = wrapHandler("cmd", handler);

    await wrapped(makeMockCmd());

    expect(flushStub.calledOnce).to.be.true;
  });
});

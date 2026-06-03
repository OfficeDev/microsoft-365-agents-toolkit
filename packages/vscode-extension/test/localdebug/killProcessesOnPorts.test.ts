// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";

import { processUtil } from "../../src/utils/processUtil";
import { killProcessesOnPorts } from "../../src/debug/depsChecker/common";
import VsCodeLogInstance from "../../src/commonlib/log";

describe("killProcessesOnPorts", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("should return 'no-pids' when no PIDs found for ports", async () => {
    sandbox.stub(processUtil, "getProcessIdsByPort").resolves([]);
    const result = await killProcessesOnPorts([3978]);
    chai.assert.equal(result, "no-pids");
  });

  it("should kill processes and return 'killed' when user confirms", async () => {
    sandbox.stub(processUtil, "getProcessIdsByPort").resolves([12345]);
    sandbox.stub(vscode.window, "showWarningMessage").resolves("Terminate Process" as any);
    const killStub = sandbox.stub(processUtil, "killProcess").resolves();

    const result = await killProcessesOnPorts([3978]);

    chai.assert.equal(result, "killed");
    chai.assert.isTrue(killStub.calledOnceWith(12345));
  });

  it("should return 'cancelled' when user dismisses notification", async () => {
    sandbox.stub(processUtil, "getProcessIdsByPort").resolves([12345]);
    sandbox.stub(vscode.window, "showWarningMessage").resolves(undefined as any);
    const killStub = sandbox.stub(processUtil, "killProcess");

    const result = await killProcessesOnPorts([3978]);

    chai.assert.equal(result, "cancelled");
    chai.assert.isTrue(killStub.notCalled);
  });

  it("should return 'copilot' when user clicks Resolve with Copilot Chat", async () => {
    sandbox.stub(processUtil, "getProcessIdsByPort").resolves([12345]);
    sandbox.stub(vscode.window, "showWarningMessage").resolves("Resolve with Copilot Chat" as any);
    const killStub = sandbox.stub(processUtil, "killProcess");

    const result = await killProcessesOnPorts([3978]);

    chai.assert.equal(result, "copilot");
    chai.assert.isTrue(killStub.notCalled);
  });

  it("should deduplicate PIDs across multiple ports", async () => {
    const getStub = sandbox.stub(processUtil, "getProcessIdsByPort");
    getStub.withArgs(3978).resolves([12345]);
    getStub.withArgs(9239).resolves([12345, 67890]);
    sandbox.stub(vscode.window, "showWarningMessage").resolves("Terminate Process" as any);
    const killStub = sandbox.stub(processUtil, "killProcess").resolves();

    const result = await killProcessesOnPorts([3978, 9239]);

    chai.assert.equal(result, "killed");
    chai.assert.equal(killStub.callCount, 2);
    const killedPids = killStub.getCalls().map((c) => c.args[0]);
    chai.assert.includeMembers(killedPids, [12345, 67890]);
  });

  it("should return 'no-pids' and log warning when an exception occurs", async () => {
    sandbox.stub(processUtil, "getProcessIdsByPort").rejects(new Error("unexpected failure"));
    const warnStub = sandbox.stub(VsCodeLogInstance, "warning");

    const result = await killProcessesOnPorts([3978]);

    chai.assert.equal(result, "no-pids");
    chai.assert.isTrue(warnStub.calledOnce);
    chai.assert.include(warnStub.firstCall.args[0], "unexpected failure");
  });

  it("should return 'no-pids' and log warning when killProcess throws", async () => {
    sandbox.stub(processUtil, "getProcessIdsByPort").resolves([12345]);
    sandbox.stub(vscode.window, "showWarningMessage").resolves("Terminate Process" as any);
    sandbox.stub(processUtil, "killProcess").rejects(new Error("kill failed"));
    const warnStub = sandbox.stub(VsCodeLogInstance, "warning");

    const result = await killProcessesOnPorts([3978]);

    chai.assert.equal(result, "no-pids");
    chai.assert.isTrue(warnStub.calledOnce);
  });

  it("should log port conflict details to output channel", async () => {
    sandbox.stub(processUtil, "getProcessIdsByPort").resolves([12345]);
    sandbox.stub(vscode.window, "showWarningMessage").resolves("Terminate Process" as any);
    sandbox.stub(processUtil, "killProcess").resolves();
    const appendStub = sandbox.stub(VsCodeLogInstance.outputChannel, "appendLine");

    await killProcessesOnPorts([3978]);

    const logCall = appendStub
      .getCalls()
      .find((c) => (c.args[0] as string).includes("[Port Conflict]"));
    chai.assert.isDefined(logCall);
    chai.assert.include(logCall!.args[0] as string, "3978");
    chai.assert.include(logCall!.args[0] as string, "12345");
  });
});

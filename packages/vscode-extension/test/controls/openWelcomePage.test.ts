import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";
import * as globalState from "@microsoft/teamsfx-core/build/common/globalState";
import { openWelcomePageAfterExtensionInstallation } from "../../src/controls/openWelcomePage";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";

describe("openWelcomePageAfterExtensionInstallation()", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("will not open welcome page if shown before", async () => {
    sandbox.stub(globalState, "globalStateGet").resolves(true);
    const globalStateUpdateStub = sandbox.stub(globalState, "globalStateUpdate");

    await openWelcomePageAfterExtensionInstallation();

    chai.assert.isTrue(globalStateUpdateStub.notCalled);
  });

  it("opens welcome page if not shown before", async () => {
    sandbox.stub(globalState, "globalStateGet").resolves(false);
    const globalStateUpdateStub = sandbox.stub(globalState, "globalStateUpdate");
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();

    await openWelcomePageAfterExtensionInstallation();

    chai.assert.isTrue(globalStateUpdateStub.calledOnce);
    chai.assert.isTrue(executeCommandStub.calledThrice);
  });
});

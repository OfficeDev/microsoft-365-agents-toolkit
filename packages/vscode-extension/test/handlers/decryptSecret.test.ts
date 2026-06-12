import * as sinon from "sinon";
import * as vscode from "vscode";
import * as globalVariables from "../../src/globalVariables";
import * as vsc_ui from "../../src/qm/vsc_ui";
import { ok, err, UserError } from "@microsoft/teamsfx-api";
import { decryptSecret } from "../../src/handlers/decryptSecret";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { MockCore } from "../mocks/mockCore";

describe("decryptSecret", function () {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("successfully update secret", async () => {
    sandbox.stub(globalVariables, "context").value({ extensionPath: "" });
    sandbox.stub(globalVariables, "core").value(new MockCore());
    const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    const sendTelemetryErrorEvent = sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
    const decrypt = sandbox.spy(globalVariables.core, "decrypt");
    const encrypt = sandbox.spy(globalVariables.core, "encrypt");
    sandbox.stub(vscode.commands, "executeCommand");
    const editBuilder = sandbox.spy();
    sandbox.stub(vscode.window, "activeTextEditor").value({
      edit: function (callback: (eb: any) => void) {
        callback({
          replace: editBuilder,
        });
      },
    });
    sandbox.stub(vsc_ui, "VS_CODE_UI").value({
      inputText: () => Promise.resolve(ok({ type: "success", result: "inputValue" })),
    });
    const range = new vscode.Range(new vscode.Position(0, 10), new vscode.Position(0, 15));

    await decryptSecret("test", range);

    sinon.assert.calledOnce(decrypt);
    sinon.assert.calledOnce(encrypt);
    sinon.assert.calledOnce(editBuilder);
    sinon.assert.calledTwice(sendTelemetryEvent);
    sinon.assert.notCalled(sendTelemetryErrorEvent);
  });

  it("no active editor", async () => {
    sandbox.stub(globalVariables, "context").value({ extensionPath: "" });
    sandbox.stub(globalVariables, "core").value(new MockCore());
    const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    const decrypt = sandbox.stub(globalVariables.core, "decrypt");
    sandbox.stub(vscode.commands, "executeCommand");
    sandbox.stub(vscode.window, "activeTextEditor");
    const range = new vscode.Range(new vscode.Position(0, 10), new vscode.Position(0, 15));

    await decryptSecret("test", range);

    sinon.assert.notCalled(decrypt);
    sinon.assert.calledOnce(sendTelemetryEvent);
  });

  it("failed to update due to corrupted secret", async () => {
    sandbox.stub(globalVariables, "context").value({ extensionPath: "" });
    sandbox.stub(globalVariables, "core").value(new MockCore());
    const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    const sendTelemetryErrorEvent = sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
    const decrypt = sandbox.stub(globalVariables.core, "decrypt");
    decrypt.returns(Promise.resolve(err(new UserError("", "fake error", ""))));
    const encrypt = sandbox.spy(globalVariables.core, "encrypt");
    sandbox.stub(vscode.commands, "executeCommand");
    const editBuilder = sandbox.spy();
    sandbox.stub(vscode.window, "activeTextEditor").value({
      edit: function (callback: (eb: any) => void) {
        callback({
          replace: editBuilder,
        });
      },
    });
    const showMessage = sandbox.stub(vscode.window, "showErrorMessage").resolves(undefined);
    const range = new vscode.Range(new vscode.Position(0, 10), new vscode.Position(0, 15));

    await decryptSecret("test", range);

    sinon.assert.calledOnce(decrypt);
    sinon.assert.notCalled(encrypt);
    sinon.assert.notCalled(editBuilder);
    sinon.assert.calledOnce(showMessage);
    sinon.assert.calledOnce(sendTelemetryEvent);
    sinon.assert.calledOnce(sendTelemetryErrorEvent);
  });

  it("failed to encrypt secret", async () => {
    sandbox.stub(globalVariables, "context").value({ extensionPath: "" });
    sandbox.stub(globalVariables, "core").value(new MockCore());
    const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    const sendTelemetryErrorEvent = sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
    const decrypt = sandbox.spy(globalVariables.core, "decrypt");
    const encrypt = sandbox
      .stub(globalVariables.core, "encrypt")
      .resolves(err(new UserError("", "fake error", "")));
    sandbox.stub(vscode.commands, "executeCommand");
    const editBuilder = sandbox.spy();
    sandbox.stub(vscode.window, "activeTextEditor").value({
      edit: function (callback: (eb: any) => void) {
        callback({
          replace: editBuilder,
        });
      },
    });
    sandbox.stub(vsc_ui, "VS_CODE_UI").value({
      inputText: () => Promise.resolve(ok({ type: "success", result: "inputValue" })),
    });
    const range = new vscode.Range(new vscode.Position(0, 10), new vscode.Position(0, 15));

    await decryptSecret("test", range);

    sinon.assert.calledOnce(decrypt);
    sinon.assert.calledOnce(encrypt);
    sinon.assert.notCalled(editBuilder);
    sinon.assert.calledOnce(sendTelemetryEvent);
    sinon.assert.calledOnce(sendTelemetryErrorEvent);
    sinon.assert.match(sendTelemetryErrorEvent.getCall(0).args[0], "edit-secret");
  });
});

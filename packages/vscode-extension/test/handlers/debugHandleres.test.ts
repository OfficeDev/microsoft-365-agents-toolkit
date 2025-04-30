import { Inputs, err, ok } from "@microsoft/teamsfx-api";
import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";
import * as launch from "../../src/debug/launch";
import * as runIconHandler from "../../src/debug/runIconHandler";
import * as globalVariables from "../../src/globalVariables";
import {
  debugInTestToolHandler,
  selectAndDebugHandler,
  treeViewLocalDebugHandler,
  treeViewPreviewHandler,
} from "../../src/handlers/debugHandlers";
import * as sharedOpts from "../../src/handlers/sharedOpts";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { TelemetryEvent } from "../../src/telemetry/extTelemetryEvents";
import * as localizeUtils from "../../src/utils/localizeUtils";
import * as systemEnvUtils from "../../src/utils/systemEnvUtils";
import { MockCore } from "../mocks/mockCore";

describe("DebugHandlers", () => {
  describe("DebugInTestTool", () => {
    const sandbox = sinon.createSandbox();

    afterEach(async () => {
      sandbox.restore();
    });

    it("treeViewDebugInTestToolHandler", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");

      await debugInTestToolHandler("treeview")();

      chai.assert.isTrue(
        executeCommandStub.calledOnceWith(
          "workbench.action.quickOpen",
          "debug Debug in Microsoft 365 Agents Playground"
        )
      );
    });

    it("messageDebugInTestToolHandler", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");

      await debugInTestToolHandler("message")();

      chai.assert.isTrue(
        executeCommandStub.calledOnceWith(
          "workbench.action.quickOpen",
          "debug Debug in Microsoft 365 Agents Playground"
        )
      );
    });
  });

  describe("TreeViewPreviewHandler", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("treeViewPreviewHandler() - previewWithManifest error", async () => {
      sandbox.stub(localizeUtils, "localize").returns("");
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sandbox.stub(systemEnvUtils, "getSystemInputs").returns({} as Inputs);
      sandbox.stub(globalVariables, "core").value(new MockCore());
      sandbox
        .stub(globalVariables.core, "previewWithManifest")
        .resolves(err({ foo: "bar" } as any));

      const result = await treeViewPreviewHandler("dev");

      chai.assert.isTrue(result.isErr());
    });

    it("treeViewPreviewHandler() - happy path", async () => {
      sandbox.stub(localizeUtils, "localize").returns("");
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sandbox.stub(systemEnvUtils, "getSystemInputs").returns({} as Inputs);
      sandbox.stub(globalVariables, "core").value(new MockCore());
      sandbox.stub(globalVariables.core, "previewWithManifest").resolves(ok("test-url"));
      sandbox.stub(launch, "openHubWebClient").resolves();

      const result = await treeViewPreviewHandler("dev");

      chai.assert.isTrue(result.isOk());
    });
  });

  describe("SelectAndDebugHandler", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("Happy path", async () => {
      const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      const selectAndDebugStub = sandbox.stub(runIconHandler, "selectAndDebug").resolves(ok(null));
      const processResultStub = sandbox.stub(sharedOpts, "processResult");

      await selectAndDebugHandler();

      chai.assert.isTrue(sendTelemetryEventStub.calledOnce);
      chai.assert.equal(
        sendTelemetryEventStub.getCall(0).args[0],
        TelemetryEvent.RunIconDebugStart
      );
      chai.assert.isTrue(selectAndDebugStub.calledOnce);
      chai.assert.isTrue(processResultStub.calledOnce);
      chai.assert.equal(processResultStub.getCall(0).args[0], TelemetryEvent.RunIconDebug);
    });
  });

  describe("TreeViewLocalDebugHandler", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("Happy path", async () => {
      const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");

      await treeViewLocalDebugHandler();

      chai.assert.isTrue(sendTelemetryEventStub.calledOnceWith(TelemetryEvent.TreeViewLocalDebug));
      chai.assert.isTrue(executeCommandStub.calledOnceWith("workbench.action.quickOpen", "debug "));
    });
  });
});

import { Inputs, err, ok } from "@microsoft/teamsfx-api";
import * as chai from "chai";
import * as vscode from "vscode";
import * as globalVariables from "../../src/globalVariables";
import { vi } from "vitest";
import { mockValue } from "../mocks/vitestMockUtils";
import {
  debugHandlersDeps,
  debugInTestToolHandler,
  selectAndDebugHandler,
  treeViewLocalDebugHandler,
  treeViewPreviewHandler,
} from "../../src/handlers/debugHandlers";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { TelemetryEvent } from "../../src/telemetry/extTelemetryEvents";
import * as localizeUtils from "../../src/utils/localizeUtils";
import { MockCore } from "../mocks/mockCore";

describe("DebugHandlers", () => {
  describe("DebugInTestTool", () => {
    afterEach(async () => {
      vi.restoreAllMocks();
    });

    it("treeViewDebugInTestToolHandler", async () => {
      mockValue(globalVariables, "core", new MockCore());
      vi.spyOn(ExtTelemetry, "sendTelemetryEvent");
      const executeCommandStub = vi.spyOn(vscode.commands, "executeCommand");

      await debugInTestToolHandler("treeview")();

      chai.assert.isTrue(
        executeCommandStub.calledOnceWith(
          "workbench.action.quickOpen",
          "debug Debug in Microsoft 365 Agents Playground"
        )
      );
    });

    it("messageDebugInTestToolHandler", async () => {
      mockValue(globalVariables, "core", new MockCore());
      vi.spyOn(ExtTelemetry, "sendTelemetryEvent");
      const executeCommandStub = vi.spyOn(vscode.commands, "executeCommand");

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
    it("treeViewPreviewHandler() - previewWithManifest error", async () => {
      vi.spyOn(localizeUtils, "localize").mockReturnValue("");
      vi.spyOn(ExtTelemetry, "sendTelemetryEvent");
      vi.spyOn(ExtTelemetry, "sendTelemetryErrorEvent");
      vi.spyOn(debugHandlersDeps, "getSystemInputs").mockReturnValue({} as Inputs);
      mockValue(globalVariables, "core", new MockCore());
      vi.spyOn(globalVariables.core, "previewWithManifest").mockResolvedValue(
        err({ foo: "bar" } as any)
      );

      const result = await treeViewPreviewHandler("dev");

      chai.assert.isTrue(result.isErr());
    });

    it("treeViewPreviewHandler() - happy path", async () => {
      vi.spyOn(localizeUtils, "localize").mockReturnValue("");
      vi.spyOn(ExtTelemetry, "sendTelemetryEvent");
      vi.spyOn(ExtTelemetry, "sendTelemetryErrorEvent");
      vi.spyOn(debugHandlersDeps, "getSystemInputs").mockReturnValue({} as Inputs);
      mockValue(globalVariables, "core", new MockCore());
      vi.spyOn(globalVariables.core, "previewWithManifest").mockResolvedValue(ok("test-url"));
      vi.spyOn(debugHandlersDeps, "openHubWebClient").mockResolvedValue();

      const result = await treeViewPreviewHandler("dev");

      chai.assert.isTrue(result.isOk());
    });
  });

  describe("SelectAndDebugHandler", () => {
    it("Happy path", async () => {
      const sendTelemetryEventStub = vi.spyOn(ExtTelemetry, "sendTelemetryEvent");
      const selectAndDebugStub = vi
        .spyOn(debugHandlersDeps, "selectAndDebug")
        .mockResolvedValue(ok(null));
      const processResultStub = vi.spyOn(debugHandlersDeps, "processResult");

      await selectAndDebugHandler();

      chai.assert.isTrue(sendTelemetryEventStub.called);
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
    it("Happy path", async () => {
      const sendTelemetryEventStub = vi.spyOn(ExtTelemetry, "sendTelemetryEvent");
      const executeCommandStub = vi.spyOn(vscode.commands, "executeCommand");

      await treeViewLocalDebugHandler();

      chai.assert.isTrue(sendTelemetryEventStub.calledOnceWith(TelemetryEvent.TreeViewLocalDebug));
      chai.assert.isTrue(executeCommandStub.calledOnceWith("workbench.action.quickOpen", "debug "));
    });
  });
});

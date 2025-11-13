import { err, FxError, Inputs, ok, Result, Stage, UserError } from "@microsoft/teamsfx-api";
import { QuestionNames, UserCancelError } from "@microsoft/teamsfx-core";
import { assert } from "chai";
import fs from "fs-extra";
import * as sinon from "sinon";
import * as vscode from "vscode";
import * as globalVariables from "../../src/globalVariables";
import {
  buildPackageHandler,
  publishInDeveloperPortalHandler,
  syncManifestHandler,
  updatePreviewManifest,
  validateManifestHandler,
} from "../../src/handlers/manifestHandlers";
import * as shared from "../../src/handlers/sharedOpts";
import * as vsc_ui from "../../src/qm/vsc_ui";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { MockCore } from "../mocks/mockCore";
describe("Manifest handlers", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
  });

  afterEach(() => {
    sandbox.restore();
  });
  describe("validateManifestHandler", () => {
    it("happy", async () => {
      sandbox.stub(shared, "runCommand").resolves(ok(undefined));
      const res = await validateManifestHandler();
      assert.isTrue(res.isOk());
    });
  });
  describe("buildPackageHandler", function () {
    it("happy()", async () => {
      sandbox.stub(shared, "runCommand").resolves(ok(undefined));
      const res = await buildPackageHandler();
      assert.isTrue(res.isOk());
    });
  });
  describe("publishInDeveloperPortalHandler", async () => {
    beforeEach(() => {
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("path"));
    });
    it("publish in developer portal - success", async () => {
      sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI({} as vscode.ExtensionContext));
      sandbox
        .stub(vsc_ui.VS_CODE_UI, "selectFile")
        .resolves(ok({ type: "success", result: "test.zip" }));
      sandbox.stub(shared, "runCommand").resolves(ok(undefined));
      sandbox
        .stub(vsc_ui.VS_CODE_UI, "selectOption")
        .resolves(ok({ type: "success", result: "test.zip" }));
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readdir").resolves(["test.zip", "test.json"] as any);
      sandbox.stub(fs, "existsSync").returns(true);
      const res = await publishInDeveloperPortalHandler();
      assert.isTrue(res.isOk());
      const res2 = await publishInDeveloperPortalHandler();
      assert.isTrue(res2.isOk());
    });

    it("publish in developer portal - cancelled", async () => {
      sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI({} as vscode.ExtensionContext));
      sandbox
        .stub(vsc_ui.VS_CODE_UI, "selectFile")
        .resolves(ok({ type: "success", result: "test2.zip" }));
      sandbox.stub(vsc_ui.VS_CODE_UI, "selectOption").resolves(err(new UserCancelError("VSC")));
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readdir").resolves(["test.zip", "test.json"] as any);
      const res = await publishInDeveloperPortalHandler();
      assert.isTrue(res.isOk());
    });
    it("select file error", async () => {
      sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI({} as vscode.ExtensionContext));
      sandbox.stub(vsc_ui.VS_CODE_UI, "selectFile").resolves(err(new UserCancelError("VSC")));
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readdir").resolves(["test.zip", "test.json"] as any);
      const res = await publishInDeveloperPortalHandler();
      assert.isTrue(res.isOk());
    });
    it("runCommand error", async () => {
      sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI({} as vscode.ExtensionContext));
      sandbox
        .stub(vsc_ui.VS_CODE_UI, "selectFile")
        .resolves(ok({ type: "success", result: "test.zip" }));
      sandbox.stub(shared, "runCommand").resolves(err(new UserCancelError("VSC")));
      sandbox
        .stub(vsc_ui.VS_CODE_UI, "selectOption")
        .resolves(ok({ type: "success", result: "test.zip" }));
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readdir").resolves(["test.zip", "test.json"] as any);
      const res = await publishInDeveloperPortalHandler();
      assert.isTrue(res.isErr());
    });
  });

  describe("updatePreviewManifest", () => {
    it("happy", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const openTextDocumentStub = sandbox
        .stub(vscode.workspace, "openTextDocument")
        .returns(Promise.resolve("" as any));
      sandbox.stub(shared, "runCommand").resolves(ok(undefined));
      await updatePreviewManifest([]);
      assert.isTrue(openTextDocumentStub.calledOnce);
    });
    it("getSelectedEnv error", async () => {
      const core = new MockCore();
      sandbox.stub(globalVariables, "core").value(core);
      sandbox.stub(core, "getSelectedEnv").resolves(err(new UserCancelError("VSC")));
      sandbox.stub(shared, "runCommand").resolves(ok(undefined));
      const res = await updatePreviewManifest([]);
      assert.isTrue(res.isErr());
    });
  });
  describe("syncManifest", () => {
    it("happy", async () => {
      const runCommandStub = sandbox.stub(shared, "runCommand").resolves(ok(undefined));
      await syncManifestHandler();
      assert.isTrue(runCommandStub.calledOnce);
    });
    it("teams app id in the input", async () => {
      const runCommandStub = sandbox
        .stub(shared, "runCommand")
        .callsFake((stage: Stage, inputs: Inputs | undefined): Promise<Result<any, FxError>> => {
          if (inputs && inputs[QuestionNames.TeamsAppId] === "teamsAppId") {
            return Promise.resolve(ok(undefined));
          }
          return Promise.resolve(err(new UserError("ut", "error", "", "")));
        });
      const res = await syncManifestHandler("teamsAppId");
      assert.isTrue(runCommandStub.calledOnce);
      assert.isTrue(res.isOk());
    });
  });
});

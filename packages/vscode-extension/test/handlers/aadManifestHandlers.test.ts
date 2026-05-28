import * as sinon from "sinon";
import * as chai from "chai";
import fs from "fs-extra";
import * as globalVariables from "../../src/globalVariables";
import * as vsc_ui from "../../src/qm/vsc_ui";
import * as vscode from "vscode";
import * as projectSettingsHelper from "@microsoft/teamsfx-core/build/common/projectSettingsHelper";
import * as localizeUtils from "@microsoft/teamsfx-core/build/common/localizeUtils";
import * as errorCommon from "../../src/error/common";
import * as sharedOpts from "../../src/handlers/sharedOpts";
import * as envHandlers from "../../src/handlers/envHandlers";
import { FxError, err, ok } from "@microsoft/teamsfx-api";
import { environmentManager } from "@microsoft/teamsfx-core";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { MockCore } from "../mocks/mockCore";
import {
  editAadManifestTemplateHandler,
  openPreviewAadFileHandler,
  updateAadAppManifestHandler,
} from "../../src/handlers/aadManifestHandlers";

describe("aadManifestHandlers", () => {
  describe("openPreviewAadFileHandler", () => {
    const sandbox = sinon.createSandbox();

    beforeEach(() => {
      sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("project is not valid", async () => {
      const core = new MockCore();
      sandbox.stub(globalVariables, "core").value(core);
      sandbox.stub(projectSettingsHelper, "isValidProject").returns(false);
      sandbox.stub(localizeUtils, "getDefaultString").returns("InvalidProjectError");
      sandbox.stub(localizeUtils, "getLocalizedString").returns("InvalidProjectError");
      const res = await openPreviewAadFileHandler([]);
      chai.assert.isTrue(res.isErr());
      chai.assert.equal(res.isErr() ? res.error.message : "Not Err", "InvalidProjectError");
    });

    it("select Env returns error", async () => {
      const core = new MockCore();
      sandbox.stub(globalVariables, "core").value(core);
      sandbox.stub(projectSettingsHelper, "isValidProject").returns(true);
      sandbox.stub(envHandlers, "askTargetEnvironment").resolves(err("selectEnvErr") as any);
      const res = await openPreviewAadFileHandler([]);
      chai.assert.isTrue(res.isErr());
      chai.assert.equal(res.isErr() ? res.error : "Not Err", "selectEnvErr");
    });

    it("runCommand returns error", async () => {
      const core = new MockCore();
      sandbox.stub(globalVariables, "core").value(core);
      sandbox.stub(projectSettingsHelper, "isValidProject").returns(true);
      sandbox.stub(envHandlers, "askTargetEnvironment").resolves(ok("dev"));
      sandbox.stub(sharedOpts, "runCommand").resolves(err("runCommandErr") as any);
      const res = await openPreviewAadFileHandler([]);
      chai.assert.isTrue(res.isErr());
      chai.assert.equal(res.isErr() ? res.error : "Not Err", "runCommandErr");
    });

    it("manifest file not exists", async () => {
      const core = new MockCore();
      sandbox.stub(globalVariables, "core").value(core);
      sandbox.stub(projectSettingsHelper, "isValidProject").returns(true);
      sandbox.stub(fs, "existsSync").returns(false);
      sandbox.stub(environmentManager, "listAllEnvConfigs").resolves(ok(["dev"]));
      sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
      sandbox.stub(vsc_ui.VS_CODE_UI, "selectOption").resolves(
        ok({
          type: "success",
          result: "dev",
        })
      );
      sandbox.stub(envHandlers, "askTargetEnvironment").resolves(ok("dev"));
      sandbox.stub(errorCommon, "showError").callsFake(async () => {});
      sandbox.stub(globalVariables.core, "buildAadManifest").resolves(ok(undefined));
      const res = await openPreviewAadFileHandler([]);
      chai.assert.isTrue(res.isErr());
    });

    it("happy path", async () => {
      const core = new MockCore();
      sandbox.stub(globalVariables, "core").value(core);
      sandbox.stub(projectSettingsHelper, "isValidProject").returns(true);
      sandbox.stub(fs, "existsSync").returns(true);
      sandbox.stub(environmentManager, "listAllEnvConfigs").resolves(ok(["dev"]));
      sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
      sandbox.stub(vsc_ui.VS_CODE_UI, "selectOption").resolves(
        ok({
          type: "success",
          result: "dev",
        })
      );
      sandbox.stub(envHandlers, "askTargetEnvironment").resolves(ok("dev"));
      sandbox.stub(errorCommon, "showError").callsFake(async () => {});
      sandbox.stub(globalVariables.core, "buildAadManifest").resolves(ok(undefined));
      sandbox.stub(vscode.workspace, "openTextDocument").resolves();
      sandbox.stub(vscode.window, "showTextDocument").resolves();

      const res = await openPreviewAadFileHandler([]);
      chai.assert.isTrue(res.isOk());
    });
  });

  describe("updateAadAppManifestHandler", () => {
    const sandbox = sinon.createSandbox();

    beforeEach(() => {
      sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("deployAadAppmanifest", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const deployAadManifest = sandbox.spy(globalVariables.core, "deployAadManifest");
      await updateAadAppManifestHandler([{ fsPath: "path/aad.dev.template" }]);
      sandbox.assert.calledOnce(deployAadManifest);
    });
  });

  describe("editAadManifestTemplate", () => {
    const sandbox = sinon.createSandbox();

    beforeEach(() => {
      sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("happy path", async () => {
      const workspacePath = "/test/workspace/path";
      const workspaceUri = vscode.Uri.file(workspacePath);
      sandbox.stub(globalVariables, "workspaceUri").value(workspaceUri);

      const openTextDocumentStub = sandbox
        .stub(vscode.workspace, "openTextDocument")
        .resolves({} as any);
      sandbox.stub(vscode.window, "showTextDocument");

      await editAadManifestTemplateHandler([null, "testTrigger"]);

      sandbox.assert.calledOnceWithExactly(
        openTextDocumentStub as any,
        `${workspaceUri.fsPath}/aad.manifest.json`
      );
    });

    it("happy path: no parameter", async () => {
      const workspacePath = "/test/workspace/path";
      const workspaceUri = vscode.Uri.file(workspacePath);
      sandbox.stub(globalVariables, "workspaceUri").value(workspaceUri);

      sandbox.stub(vscode.workspace, "openTextDocument").resolves({} as any);
      const showTextDocumentStub = sandbox.stub(vscode.window, "showTextDocument");

      await editAadManifestTemplateHandler([]);

      chai.assert.isTrue(showTextDocumentStub.callCount === 0);
    });

    it("happy path: workspaceUri is undefined", async () => {
      const workspaceUri = undefined;
      sandbox.stub(globalVariables, "workspaceUri").value(undefined);

      const openTextDocumentStub = sandbox
        .stub(vscode.workspace, "openTextDocument")
        .resolves({} as any);
      sandbox.stub(vscode.window, "showTextDocument");

      await editAadManifestTemplateHandler([null, "testTrigger"]);

      sandbox.assert.calledOnceWithExactly(
        openTextDocumentStub as any,
        `${workspaceUri}/aad.manifest.json`
      );
    });
  });
});

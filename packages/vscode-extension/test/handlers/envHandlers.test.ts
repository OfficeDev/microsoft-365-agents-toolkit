import { ConfigFolderName, err, ok, Void } from "@microsoft/teamsfx-api";
import { environmentManager, pathUtils } from "@microsoft/teamsfx-core";
import * as localizeUtils from "@microsoft/teamsfx-core/build/common/localizeUtils";
import * as projectSettingsHelper from "@microsoft/teamsfx-core/build/common/projectSettingsHelper";
import * as chai from "chai";
import fs from "fs-extra";
import path from "path";
import * as sinon from "sinon";
import * as vscode from "vscode";
import { ExtensionErrors } from "../../src/error/error";
import * as globalVariables from "../../src/globalVariables";
import {
  askTargetEnvironment,
  createNewEnvironment,
  envHandlersDeps,
  openConfigStateFile,
  refreshEnvironment,
} from "../../src/handlers/envHandlers";
import * as shared from "../../src/handlers/sharedOpts";
import * as vsc_ui from "../../src/qm/vsc_ui";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import envTreeProviderInstance from "../../src/treeview/environmentTreeViewProvider";

describe("Env handlers", () => {
  describe("createNewEnvironment", () => {
    const sandbox = sinon.createSandbox();

    beforeEach(() => {
      sandbox.stub(envHandlersDeps, "sendTelemetryEvent");
      sandbox.stub(envHandlersDeps, "sendTelemetryErrorEvent");
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("happy", async () => {
      sandbox.stub(envHandlersDeps, "reloadEnvironments").resolves(ok(Void));
      sandbox.stub(envHandlersDeps, "runCommand").resolves(ok(undefined));
      const res = await createNewEnvironment();
      chai.assert.isTrue(res.isOk());
    });
  });

  describe("refreshEnvironment", () => {
    const sandbox = sinon.createSandbox();

    beforeEach(() => {
      sandbox.stub(envHandlersDeps, "sendTelemetryEvent");
      sandbox.stub(envHandlersDeps, "sendTelemetryErrorEvent");
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("happy", async () => {
      sandbox.stub(envHandlersDeps, "reloadEnvironments").resolves(ok(Void));
      const res = await refreshEnvironment();
      chai.assert.isTrue(res.isOk());
    });
  });

  describe("openConfigStateFile", () => {
    const sandbox = sinon.createSandbox();

    beforeEach(() => {
      sandbox.stub(envHandlersDeps, "sendTelemetryEvent");
      sandbox.stub(envHandlersDeps, "sendTelemetryErrorEvent");
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("InvalidArgs", async () => {
      sandbox.stub(envHandlersDeps, "isValidProject").returns(true);
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("./tmp"));

      const res = await openConfigStateFile([]);

      if (res) {
        chai.assert.isTrue(res.isErr());
        chai.assert.equal(res.error.name, ExtensionErrors.InvalidArgs);
      }
    });

    it("noOpenWorkspace", async () => {
      sandbox.stub(envHandlersDeps, "isValidProject").returns(true);
      sandbox.stub(globalVariables, "workspaceUri").value({ fsPath: undefined });

      const res = await openConfigStateFile([]);

      if (res) {
        chai.assert.isTrue(res.isErr());
        chai.assert.equal(res.error.name, ExtensionErrors.NoWorkspaceError);
      }
    });

    it("invalidProject", async () => {
      sandbox.stub(envHandlersDeps, "isValidProject").returns(false);
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("./tmp"));

      const res = await openConfigStateFile([{ env: "dev" }]);

      if (res) {
        chai.assert.isTrue(res.isErr());
        chai.assert.equal(res.error.name, ExtensionErrors.InvalidProject);
      }
    });

    it("invalid target environment", async () => {
      sandbox.stub(envHandlersDeps, "isValidProject").returns(true);
      sandbox.stub(envHandlersDeps, "listAllEnvConfigs").resolves(ok([]));
      sandbox
        .stub(envHandlersDeps, "selectOption")
        .resolves(err({ error: "invalid target env" } as any));
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("./tmp"));

      const res = await openConfigStateFile([{ env: undefined, type: "env" }]);

      if (res) {
        chai.assert.isTrue(res.isErr());
      }
    });

    it("valid args", async () => {
      const env = "remote";
      sandbox.stub(envHandlersDeps, "isValidProject").returns(true);
      sandbox.stub(envHandlersDeps, "getEnvFolderPath").resolves(ok(env));
      sandbox.stub(envHandlersDeps, "pathExists").resolves(false);
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("./tmp"));

      const res = await openConfigStateFile([{ env: env, type: "env", from: "aad" }]);

      if (res) {
        chai.assert.isTrue(res.isErr());
        chai.assert.equal(res.error.name, ExtensionErrors.EnvFileNotFoundError);
      }
    });

    it("invalid env folder", async () => {
      sandbox.stub(envHandlersDeps, "isValidProject").returns(true);
      sandbox.stub(envHandlersDeps, "getEnvFolderPath").resolves(err({ error: "unknown" } as any));
      sandbox.stub(envHandlersDeps, "pathExists").resolves(true);
      sandbox.stub(envHandlersDeps, "openTextDocument").resolves("" as any);
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("./tmp"));

      const res = await openConfigStateFile([{ env: "local", type: "env" }]);

      if (res) {
        chai.assert.isTrue(res.isErr());
      }
    });

    it("success", async () => {
      sandbox.stub(envHandlersDeps, "isValidProject").returns(true);
      sandbox.stub(envHandlersDeps, "getEnvFolderPath").resolves(ok(""));
      sandbox.stub(envHandlersDeps, "pathExists").resolves(true);
      sandbox.stub(envHandlersDeps, "openTextDocument").resolves("" as any);
      sandbox.stub(envHandlersDeps, "showTextDocument").returns(undefined as any);
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("./tmp"));

      const res = await openConfigStateFile([{ env: "local", type: "env" }]);

      chai.assert.isTrue(res === undefined);
    });
  });

  describe("askTargetEnvironment", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("invalid project", async () => {
      sandbox.stub(envHandlersDeps, "isValidProject").returns(false);
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("./tmp"));

      const res = await askTargetEnvironment();

      chai.assert.isTrue(res.isErr());
    });

    it("success", async () => {
      sandbox.stub(envHandlersDeps, "isValidProject").returns(true);
      sandbox.stub(envHandlersDeps, "listAllEnvConfigs").resolves(ok(["dev", "prod"]));
      sandbox.stub(envHandlersDeps, "selectOption").resolves(ok({ result: "dev" } as any));
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("./tmp"));

      const res = await askTargetEnvironment();

      chai.assert.isTrue(res.isOk());
      chai.assert.equal(res.value, "dev");
    });

    it("listAllEnvConfigs returns error", async () => {
      sandbox.stub(envHandlersDeps, "isValidProject").returns(true);
      sandbox.stub(envHandlersDeps, "listAllEnvConfigs").resolves(err({ error: "unknown" } as any));
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("./tmp"));

      const res = await askTargetEnvironment();

      chai.assert.isTrue(res.isErr());
    });
  });
});

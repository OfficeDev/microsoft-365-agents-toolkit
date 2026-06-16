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
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("happy", async () => {
      sandbox.stub(envTreeProviderInstance, "reloadEnvironments").resolves(ok(Void));
      sandbox.stub(shared, "runCommand").resolves(ok(undefined));
      const res = await createNewEnvironment();
      chai.assert.isTrue(res.isOk());
    });
  });

  describe("refreshEnvironment", () => {
    const sandbox = sinon.createSandbox();

    beforeEach(() => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("happy", async () => {
      sandbox.stub(envTreeProviderInstance, "reloadEnvironments").resolves(ok(Void));
      const res = await refreshEnvironment();
      chai.assert.isTrue(res.isOk());
    });
  });

  describe("openConfigStateFile", () => {
    const sandbox = sinon.createSandbox();

    beforeEach(() => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("InvalidArgs", async () => {
      const env = "local";
      const tmpDir = fs.mkdtempSync(path.resolve("./tmp"));

      sandbox.stub(projectSettingsHelper, "isValidProject").returns(true);
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file(tmpDir));
      const projectSettings: any = {
        appName: "myapp",
        version: "1.0.0",
        projectId: "123",
      };
      const configFolder = path.resolve(tmpDir, `.${ConfigFolderName}`, "configs");
      await fs.mkdir(configFolder, { recursive: true });
      const settingsFile = path.resolve(configFolder, "projectSettings.json");
      await fs.writeJSON(settingsFile, JSON.stringify(projectSettings, null, 4));

      sandbox.stub(globalVariables, "context").value({ extensionPath: path.resolve("../../") });
      sandbox.stub(vsc_ui, "VS_CODE_UI").value({
        selectOption: () => Promise.resolve(ok({ type: "success", result: env })),
      });

      const res = await openConfigStateFile([]);
      await fs.remove(tmpDir);

      if (res) {
        chai.assert.isTrue(res.isErr());
        chai.assert.equal(res.error.name, ExtensionErrors.InvalidArgs);
      }
    });

    it("noOpenWorkspace", async () => {
      const env = "local";

      sandbox.stub(globalVariables, "workspaceUri").value({ fsPath: undefined });

      sandbox.stub(globalVariables, "context").value({ extensionPath: path.resolve("../../") });
      sandbox.stub(vsc_ui, "VS_CODE_UI").value({
        selectOption: () => Promise.resolve(ok({ type: "success", result: env })),
      });

      const res = await openConfigStateFile([]);

      if (res) {
        chai.assert.isTrue(res.isErr());
        chai.assert.equal(res.error.name, ExtensionErrors.NoWorkspaceError);
      }
    });

    it("invalidProject", async () => {
      const env = "local";
      const tmpDir = fs.mkdtempSync(path.resolve("./tmp"));

      sandbox.stub(projectSettingsHelper, "isValidProject").returns(false);

      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file(tmpDir));
      sandbox.stub(globalVariables, "context").value({ extensionPath: path.resolve("../../") });
      sandbox.stub(vsc_ui, "VS_CODE_UI").value({
        selectOption: () => Promise.resolve(ok({ type: "success", result: env })),
      });

      const res = await openConfigStateFile([]);
      await fs.remove(tmpDir);

      if (res) {
        chai.assert.isTrue(res.isErr());
        chai.assert.equal(res.error.name, ExtensionErrors.InvalidProject);
      }
    });

    it("invalid target environment", async () => {
      const env = "local";
      const tmpDir = fs.mkdtempSync(path.resolve("./tmp"));

      sandbox.stub(projectSettingsHelper, "isValidProject").returns(true);
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file(tmpDir));
      const projectSettings: any = {
        appName: "myapp",
        version: "1.0.0",
        projectId: "123",
      };
      const configFolder = path.resolve(tmpDir, `.${ConfigFolderName}`, "configs");
      await fs.mkdir(configFolder, { recursive: true });
      const settingsFile = path.resolve(configFolder, "projectSettings.json");
      await fs.writeJSON(settingsFile, JSON.stringify(projectSettings, null, 4));

      sandbox.stub(globalVariables, "context").value({ extensionPath: path.resolve("../../") });
      sandbox.stub(vsc_ui, "VS_CODE_UI").value({
        selectOption: () => Promise.resolve(err({ error: "invalid target env" })),
      });
      sandbox.stub(environmentManager, "listAllEnvConfigs").resolves(ok([]));
      sandbox.stub(fs, "pathExists").resolves(false);
      sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok(env));

      const res = await openConfigStateFile([{ env: undefined, type: "env" }]);
      await fs.remove(tmpDir);

      if (res) {
        chai.assert.isTrue(res.isErr());
      }
    });

    it("valid args", async () => {
      const env = "remote";
      const tmpDir = fs.mkdtempSync(path.resolve("./tmp"));

      sandbox.stub(projectSettingsHelper, "isValidProject").returns(true);
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file(tmpDir));
      const projectSettings: any = {
        appName: "myapp",
        version: "1.0.0",
        projectId: "123",
      };
      const configFolder = path.resolve(tmpDir, `.${ConfigFolderName}`, "configs");
      await fs.mkdir(configFolder, { recursive: true });
      const settingsFile = path.resolve(configFolder, "projectSettings.json");
      await fs.writeJSON(settingsFile, JSON.stringify(projectSettings, null, 4));

      sandbox.stub(globalVariables, "context").value({ extensionPath: path.resolve("../../") });
      sandbox.stub(vsc_ui, "VS_CODE_UI").value({
        selectOption: () => Promise.resolve(ok({ type: "success", result: env })),
      });
      sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok(env));
      sandbox.stub(fs, "pathExists").resolves(false);
      sandbox.stub(environmentManager, "listAllEnvConfigs").resolves(ok([]));

      const res = await openConfigStateFile([{ env: undefined, type: "env", from: "aad" }]);
      await fs.remove(tmpDir);

      if (res) {
        chai.assert.isTrue(res.isErr());
        chai.assert.equal(res.error.name, ExtensionErrors.EnvFileNotFoundError);
      }
    });

    it("invalid env folder", async () => {
      const env = "local";
      const tmpDir = fs.mkdtempSync(path.resolve("./tmp"));

      sandbox.stub(projectSettingsHelper, "isValidProject").returns(true);
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file(tmpDir));
      const projectSettings: any = {
        appName: "myapp",
        version: "1.0.0",
        projectId: "123",
      };
      const configFolder = path.resolve(tmpDir, `.${ConfigFolderName}`, "configs");
      await fs.mkdir(configFolder, { recursive: true });
      const settingsFile = path.resolve(configFolder, "projectSettings.json");
      await fs.writeJSON(settingsFile, JSON.stringify(projectSettings, null, 4));

      sandbox.stub(globalVariables, "context").value({ extensionPath: path.resolve("../../") });
      sandbox.stub(vsc_ui, "VS_CODE_UI").value({
        selectOption: () => Promise.resolve(ok({ type: "success", result: env })),
      });
      sandbox.stub(pathUtils, "getEnvFolderPath").resolves(err({ error: "unknown" } as any));
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(vscode.workspace, "openTextDocument").resolves("" as any);

      const res = await openConfigStateFile([{ env: env, type: "env" }]);
      await fs.remove(tmpDir);

      if (res) {
        chai.assert.isTrue(res.isErr());
      }
    });

    it("success", async () => {
      const env = "local";
      const tmpDir = fs.mkdtempSync(path.resolve("./tmp"));

      sandbox.stub(projectSettingsHelper, "isValidProject").returns(true);
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file(tmpDir));
      const projectSettings: any = {
        appName: "myapp",
        version: "1.0.0",
        projectId: "123",
      };
      const configFolder = path.resolve(tmpDir, `.${ConfigFolderName}`, "configs");
      await fs.mkdir(configFolder, { recursive: true });
      const settingsFile = path.resolve(configFolder, "projectSettings.json");
      await fs.writeJSON(settingsFile, JSON.stringify(projectSettings, null, 4));

      sandbox.stub(globalVariables, "context").value({ extensionPath: path.resolve("../../") });
      sandbox.stub(vsc_ui, "VS_CODE_UI").value({
        selectOption: () => Promise.resolve(ok({ type: "success", result: env })),
      });
      sandbox.stub(pathUtils, "getEnvFolderPath").resolves(ok(env));
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(vscode.workspace, "openTextDocument").returns(Promise.resolve("" as any));

      const res = await openConfigStateFile([{ env: env, type: "env" }]);
      await fs.remove(tmpDir);

      if (res) {
        chai.assert.isTrue(res.isOk());
      }
    });
  });

  describe("askTargetEnvironment", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("invalid project", async () => {
      sandbox.stub(globalVariables, "workspaceUri");
      sandbox.stub(projectSettingsHelper, "isValidProject").returns(false);
      sandbox.stub(localizeUtils, "getDefaultString").returns("InvalidProjectError");
      sandbox.stub(localizeUtils, "getLocalizedString").returns("InvalidProjectError");
      const res = await askTargetEnvironment();
      chai.assert.isTrue(res.isErr());
      chai.assert.equal(res.isErr() ? res.error.message : "Not Error", "InvalidProjectError");
    });

    it("listAllEnvConfigs returns error", async () => {
      sandbox.stub(globalVariables, "workspaceUri");
      sandbox.stub(projectSettingsHelper, "isValidProject").returns(true);
      sandbox
        .stub(environmentManager, "listAllEnvConfigs")
        .resolves(err("envProfilesResultErr") as any);
      const res = await askTargetEnvironment();
      chai.assert.isTrue(res.isErr());
      chai.assert.equal(res.isErr() ? res.error : "Not Error", "envProfilesResultErr");
    });
  });
});

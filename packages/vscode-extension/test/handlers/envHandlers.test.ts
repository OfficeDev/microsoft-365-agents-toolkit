import { ConfigFolderName, err, ok, Void } from "@microsoft/teamsfx-api";
import { environmentManager, pathUtils } from "@microsoft/teamsfx-core";
import * as localizeUtils from "@microsoft/teamsfx-core/build/common/localizeUtils";
import * as projectSettingsHelper from "@microsoft/teamsfx-core/build/common/projectSettingsHelper";
import * as chai from "chai";
import fs from "fs-extra";
import path from "path";
import * as vscode from "vscode";
import { ExtensionErrors } from "../../src/error/error";
import * as globalVariables from "../../src/globalVariables";
import { vi } from "vitest";
import { mockValue } from "../mocks/vitestMockUtils";
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
    beforeEach(() => {
      vi.spyOn(envHandlersDeps, "sendTelemetryEvent");
      vi.spyOn(envHandlersDeps, "sendTelemetryErrorEvent");
    });

    it("happy", async () => {
      vi.spyOn(envHandlersDeps, "reloadEnvironments").mockResolvedValue(ok(Void));
      vi.spyOn(envHandlersDeps, "runCommand").mockResolvedValue(ok(undefined));
      const res = await createNewEnvironment();
      chai.assert.isTrue(res.isOk());
    });
  });

  describe("refreshEnvironment", () => {
    beforeEach(() => {
      vi.spyOn(envHandlersDeps, "sendTelemetryEvent");
      vi.spyOn(envHandlersDeps, "sendTelemetryErrorEvent");
    });

    it("happy", async () => {
      vi.spyOn(envHandlersDeps, "reloadEnvironments").mockResolvedValue(ok(Void));
      const res = await refreshEnvironment();
      chai.assert.isTrue(res.isOk());
    });
  });

  describe("openConfigStateFile", () => {
    beforeEach(() => {
      vi.spyOn(envHandlersDeps, "sendTelemetryEvent");
      vi.spyOn(envHandlersDeps, "sendTelemetryErrorEvent");
    });

    it("InvalidArgs", async () => {
      vi.spyOn(envHandlersDeps, "isValidProject").mockReturnValue(true);
      mockValue(globalVariables, "workspaceUri", vscode.Uri.file("./tmp"));

      const res = await openConfigStateFile([]);

      if (res) {
        chai.assert.isTrue(res.isErr());
        chai.assert.equal(res.error.name, ExtensionErrors.InvalidArgs);
      }
    });

    it("noOpenWorkspace", async () => {
      vi.spyOn(envHandlersDeps, "isValidProject").mockReturnValue(true);
      mockValue(globalVariables, "workspaceUri", { fsPath: undefined });

      const res = await openConfigStateFile([]);

      if (res) {
        chai.assert.isTrue(res.isErr());
        chai.assert.equal(res.error.name, ExtensionErrors.NoWorkspaceError);
      }
    });

    it("invalidProject", async () => {
      vi.spyOn(envHandlersDeps, "isValidProject").mockReturnValue(false);
      mockValue(globalVariables, "workspaceUri", vscode.Uri.file("./tmp"));

      const res = await openConfigStateFile([{ env: "dev" }]);

      if (res) {
        chai.assert.isTrue(res.isErr());
        chai.assert.equal(res.error.name, ExtensionErrors.InvalidProject);
      }
    });

    it("invalid target environment", async () => {
      vi.spyOn(envHandlersDeps, "isValidProject").mockReturnValue(true);
      vi.spyOn(envHandlersDeps, "listAllEnvConfigs").mockResolvedValue(ok([]));
      vi.spyOn(envHandlersDeps, "selectOption").mockResolvedValue(
        err({ error: "invalid target env" } as any)
      );
      mockValue(globalVariables, "workspaceUri", vscode.Uri.file("./tmp"));

      const res = await openConfigStateFile([{ env: undefined, type: "env" }]);

      if (res) {
        chai.assert.isTrue(res.isErr());
      }
    });

    it("valid args", async () => {
      const env = "remote";
      vi.spyOn(envHandlersDeps, "isValidProject").mockReturnValue(true);
      vi.spyOn(envHandlersDeps, "getEnvFolderPath").mockResolvedValue(ok(env));
      vi.spyOn(envHandlersDeps, "pathExists").mockResolvedValue(false);
      mockValue(globalVariables, "workspaceUri", vscode.Uri.file("./tmp"));

      const res = await openConfigStateFile([{ env: env, type: "env", from: "aad" }]);

      if (res) {
        chai.assert.isTrue(res.isErr());
        chai.assert.equal(res.error.name, ExtensionErrors.EnvFileNotFoundError);
      }
    });

    it("invalid env folder", async () => {
      vi.spyOn(envHandlersDeps, "isValidProject").mockReturnValue(true);
      vi.spyOn(envHandlersDeps, "getEnvFolderPath").mockResolvedValue(
        err({ error: "unknown" } as any)
      );
      vi.spyOn(envHandlersDeps, "pathExists").mockResolvedValue(true);
      vi.spyOn(envHandlersDeps, "openTextDocument").mockResolvedValue("" as any);
      mockValue(globalVariables, "workspaceUri", vscode.Uri.file("./tmp"));

      const res = await openConfigStateFile([{ env: "local", type: "env" }]);

      if (res) {
        chai.assert.isTrue(res.isErr());
      }
    });

    it("success", async () => {
      vi.spyOn(envHandlersDeps, "isValidProject").mockReturnValue(true);
      vi.spyOn(envHandlersDeps, "getEnvFolderPath").mockResolvedValue(ok(""));
      vi.spyOn(envHandlersDeps, "pathExists").mockResolvedValue(true);
      vi.spyOn(envHandlersDeps, "openTextDocument").mockResolvedValue("" as any);
      vi.spyOn(envHandlersDeps, "showTextDocument").mockReturnValue(undefined as any);
      mockValue(globalVariables, "workspaceUri", vscode.Uri.file("./tmp"));

      const res = await openConfigStateFile([{ env: "local", type: "env" }]);

      chai.assert.isTrue(res === undefined);
    });
  });

  describe("askTargetEnvironment", () => {
    it("invalid project", async () => {
      vi.spyOn(envHandlersDeps, "isValidProject").mockReturnValue(false);
      mockValue(globalVariables, "workspaceUri", vscode.Uri.file("./tmp"));

      const res = await askTargetEnvironment();

      chai.assert.isTrue(res.isErr());
    });

    it("success", async () => {
      vi.spyOn(envHandlersDeps, "isValidProject").mockReturnValue(true);
      vi.spyOn(envHandlersDeps, "listAllEnvConfigs").mockResolvedValue(ok(["dev", "prod"]));
      vi.spyOn(envHandlersDeps, "selectOption").mockResolvedValue(ok({ result: "dev" } as any));
      mockValue(globalVariables, "workspaceUri", vscode.Uri.file("./tmp"));

      const res = await askTargetEnvironment();

      chai.assert.isTrue(res.isOk());
      chai.assert.equal(res.value, "dev");
    });

    it("listAllEnvConfigs returns error", async () => {
      vi.spyOn(envHandlersDeps, "isValidProject").mockReturnValue(true);
      vi.spyOn(envHandlersDeps, "listAllEnvConfigs").mockResolvedValue(
        err({ error: "unknown" } as any)
      );
      mockValue(globalVariables, "workspaceUri", vscode.Uri.file("./tmp"));

      const res = await askTargetEnvironment();

      chai.assert.isTrue(res.isErr());
    });
  });

  describe("envHandlersDeps delegation", () => {
    it("isValidProject delegates to core", () => {
      const result = envHandlersDeps.isValidProject(undefined);
      chai.expect(typeof result).to.equal("boolean");
    });

    it("listAllEnvConfigs delegates to environmentManager", async () => {
      vi.spyOn(environmentManager, "listAllEnvConfigs").mockResolvedValue(ok(["dev"]));
      const result = await envHandlersDeps.listAllEnvConfigs("/test");
      chai.expect(result.isOk()).to.be.true;
    });

    it("getEnvFolderPath delegates to pathUtils", async () => {
      vi.spyOn(pathUtils, "getEnvFolderPath").mockResolvedValue(ok("/test/.env"));
      const result = await envHandlersDeps.getEnvFolderPath("/test");
      chai.expect(result.isOk()).to.be.true;
    });

    it("pathExists delegates to fs", async () => {
      vi.spyOn(fs, "pathExists").mockResolvedValue(true as never);
      const result = await envHandlersDeps.pathExists("/test");
      chai.expect(result).to.be.true;
    });

    it("openTextDocument delegates to vscode.workspace", async () => {
      vi.spyOn(vscode.workspace, "openTextDocument").mockResolvedValue({} as any);
      await envHandlersDeps.openTextDocument("/test");
      chai.expect((vscode.workspace.openTextDocument as any).called).to.be.true;
    });

    it("showTextDocument delegates to vscode.window", async () => {
      vi.spyOn(vscode.window, "showTextDocument").mockResolvedValue({} as any);
      await envHandlersDeps.showTextDocument({} as any);
      chai.expect((vscode.window.showTextDocument as any).called).to.be.true;
    });

    it("runCommand delegates to sharedOpts.runCommand", async () => {
      vi.spyOn(shared, "runCommand").mockResolvedValue(ok(undefined));
      const result = await envHandlersDeps.runCommand("provision" as any);
      chai.expect(result.isOk()).to.be.true;
    });

    it("reloadEnvironments delegates to envTreeProviderInstance", async () => {
      vi.spyOn(envTreeProviderInstance, "reloadEnvironments").mockResolvedValue(
        ok(undefined as any)
      );
      await envHandlersDeps.reloadEnvironments();
      chai.expect((envTreeProviderInstance.reloadEnvironments as any).called).to.be.true;
    });

    it("selectOption delegates to VS_CODE_UI", async () => {
      mockValue(vsc_ui, "VS_CODE_UI", new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
      vi.spyOn(vsc_ui.VS_CODE_UI, "selectOption").mockResolvedValue(
        ok({ type: "success", result: "dev" } as any)
      );
      await envHandlersDeps.selectOption({ name: "env", title: "Select env" });
      chai.expect((vsc_ui.VS_CODE_UI.selectOption as any).called).to.be.true;
    });
  });
});

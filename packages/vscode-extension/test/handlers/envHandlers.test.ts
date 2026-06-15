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
});

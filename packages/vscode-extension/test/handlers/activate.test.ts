import * as chai from "chai";
import * as vscode from "vscode";
import path from "path";
import * as globalVariables from "../../src/globalVariables";
import { vi } from "vitest";
import { mockValue } from "../mocks/vitestMockUtils";
import {
  activate,
  activateDeps,
  refreshEnvTreeOnEnvFileChanged,
  refreshEnvTreeOnFilesNameChanged,
  refreshEnvTreeOnProjectSettingFileChanged,
} from "../../src/handlers/activate";
import { ok, signedIn, signedOut } from "@microsoft/teamsfx-api";
import { FxCore, GraphScopes } from "@microsoft/teamsfx-core";
import { FeatureFlags, featureFlagManager } from "@microsoft/teamsfx-core";
import commandController from "../../src/commandController";
import { AzureAccountManager } from "../../src/commonlib/azureLogin";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import accountTreeViewProviderInstance from "../../src/treeview/account/accountTreeViewProvider";
import envTreeProviderInstance from "../../src//treeview/environmentTreeViewProvider";
import TreeViewManagerInstance from "../../src/treeview/treeViewManager";
import M365TokenInstance from "../../src/commonlib/m365Login";
import { MockCore } from "../mocks/mockCore";

describe("Activate", function () {
  describe("activate()", function () {
    beforeEach(() => {
      vi.spyOn(accountTreeViewProviderInstance, "subscribeToStatusChanges");
      vi.spyOn(vscode.extensions, "getExtension").mockReturnValue(undefined);
      vi.spyOn(TreeViewManagerInstance, "getTreeView").mockReturnValue(undefined);
      vi.spyOn(ExtTelemetry, "dispose");
    });

    it("No globalState error", async () => {
      const result = await activate();
      chai.assert.deepEqual(result.isOk() ? result.value : result.error.name, {});
    });

    it("Valid project", async () => {
      vi.spyOn(activateDeps, "isValidProject").mockReturnValue(true);
      const sendTelemetryStub = vi.spyOn(ExtTelemetry, "sendTelemetryEvent");
      const addSharedPropertyStub = vi.spyOn(ExtTelemetry, "addSharedProperty");
      const setCommandIsRunningStub = vi.spyOn(globalVariables, "setCommandIsRunning");
      mockValue(globalVariables, "workspaceUri", vscode.Uri.parse("test"));
      const addFileSystemWatcherStub = vi.spyOn(activateDeps, "addFileSystemWatcher");
      const lockedByOperationStub = vi.spyOn(commandController, "lockedByOperation");
      const unlockedByOperationStub = vi.spyOn(commandController, "unlockedByOperation");
      const azureAccountSetStatusChangeMapStub = vi.spyOn(
        AzureAccountManager.prototype,
        "setStatusChangeMap"
      );
      const m365AccountSetStatusChangeMapStub = vi.spyOn(M365TokenInstance, "setStatusChangeMap");
      const showMessageStub = vi
        .spyOn(vscode.window, "showInformationMessage")
        .mockResolvedValue(undefined);
      let lockCallback: any;
      let unlockCallback: any;

      vi.spyOn(FxCore.prototype, "on").mockImplementation((event: string, callback: any) => {
        if (event === "lock") {
          lockCallback = callback;
        } else {
          unlockCallback = callback;
        }
      });
      azureAccountSetStatusChangeMapStub.mockImplementation(
        (
          name: string,
          statusChange: (
            status: string,
            token?: string,
            accountInfo?: Record<string, unknown>
          ) => Promise<void>
        ) => {
          statusChange(signedIn).then(() => {});
          statusChange(signedOut).then(() => {});
          return Promise.resolve(true);
        }
      );
      m365AccountSetStatusChangeMapStub.mockImplementation(
        (
          name: string,
          tokenRequest: unknown,
          statusChange: (
            status: string,
            token?: string,
            accountInfo?: Record<string, unknown>
          ) => Promise<void>
        ) => {
          statusChange(signedIn).then(() => {});
          statusChange(signedOut).then(() => {});
          return Promise.resolve(ok(true));
        }
      );
      const result = await activate();

      chai.assert.isTrue(addFileSystemWatcherStub.calledOnceWith("test"));
      chai.assert.isTrue(addSharedPropertyStub.called);
      chai.assert.isTrue(sendTelemetryStub.calledOnceWith("open-teams-app"));
      chai.assert.deepEqual(result.isOk() ? result.value : result.error.name, {});

      lockCallback("test");
      setCommandIsRunningStub.calledOnceWith(true);
      lockedByOperationStub.calledOnceWith("test");

      unlockCallback("test");
      unlockedByOperationStub.calledOnceWith("test");

      chai.assert.isTrue(showMessageStub.called);
    });

    it("uses Graph scopes for M365 status change in sovereign high", async () => {
      vi.spyOn(featureFlagManager, "getStringValue").mockReturnValue("GCC H");
      vi.spyOn(activateDeps, "isValidProject").mockReturnValue(true);
      vi.spyOn(ExtTelemetry, "addSharedProperty");
      vi.spyOn(ExtTelemetry, "sendTelemetryEvent");
      mockValue(globalVariables, "workspaceUri", vscode.Uri.parse("test"));
      vi.spyOn(activateDeps, "addFileSystemWatcher");
      vi.spyOn(commandController, "lockedByOperation");
      vi.spyOn(commandController, "unlockedByOperation");
      vi.spyOn(AzureAccountManager.prototype, "setStatusChangeMap").mockResolvedValue(true);
      const m365AccountSetStatusChangeMapStub = vi
        .spyOn(M365TokenInstance, "setStatusChangeMap")
        .mockResolvedValue(ok(true));
      vi.spyOn(vscode.window, "showInformationMessage").mockResolvedValue(undefined);
      vi.spyOn(FxCore.prototype, "on").mockReturnValue();

      const result = await activate();

      chai.assert.isTrue(result.isOk());
      chai.assert.isTrue(
        m365AccountSetStatusChangeMapStub.calledWithMatch(
          "successfully-sign-in-m365",
          { scopes: GraphScopes },
          expect.any(Function),
          false
        )
      );
    });

    it("throws error", async () => {
      vi.spyOn(activateDeps, "isValidProject").mockReturnValue(false);
      vi.spyOn(M365TokenInstance, "setStatusChangeMap");
      vi.spyOn(FxCore.prototype, "on").throws(new Error("test"));
      const showErrorMessageStub = vi
        .spyOn(vscode.window, "showErrorMessage")
        .mockResolvedValue(undefined);

      const result = await activate();

      chai.assert.isTrue(result.isErr());
      chai.assert.isTrue(showErrorMessageStub.called);
    });
  });

  describe("refreshEnvTreeOnEnvFileChanged", function () {
    it("Refresh Env", async () => {
      mockValue(globalVariables, "core", new MockCore());
      const isEnvFileStub = vi.spyOn(globalVariables.core, "isEnvFile").mockResolvedValue(ok(true));
      const reloadEnvStub = vi.spyOn(envTreeProviderInstance, "reloadEnvironments");
      await refreshEnvTreeOnEnvFileChanged("workspaceUri", [
        vscode.Uri.parse("File1"),
        vscode.Uri.parse("File2"),
      ]);
      chai.assert.isTrue(isEnvFileStub.calledOnce);
      chai.assert.isTrue(reloadEnvStub.calledOnce);
    });

    it("No need to refresh Env", async () => {
      mockValue(globalVariables, "core", new MockCore());
      const isEnvFileStub = vi
        .spyOn(globalVariables.core, "isEnvFile")
        .mockResolvedValue(ok(false));
      const reloadEnvStub = vi.spyOn(envTreeProviderInstance, "reloadEnvironments");
      await refreshEnvTreeOnEnvFileChanged("workspaceUri", [
        vscode.Uri.parse("File1"),
        vscode.Uri.parse("File2"),
      ]);
      chai.assert.isTrue(isEnvFileStub.calledTwice);
      chai.assert.isTrue(reloadEnvStub.notCalled);
    });
  });

  describe("refreshEnvTreeOnFilesNameChanged", function () {
    it("Refresh Env", async () => {
      mockValue(globalVariables, "core", new MockCore());
      const isEnvFileStub = vi
        .spyOn(globalVariables.core, "isEnvFile")
        .mockImplementation((projectPath, inputFile) => {
          if (inputFile === "File1New" || inputFile === "File2New") {
            return Promise.resolve(ok(true));
          }
          return Promise.resolve(ok(false));
        });
      const reloadEnvStub = vi.spyOn(envTreeProviderInstance, "reloadEnvironments");
      await refreshEnvTreeOnFilesNameChanged("workspaceUri", {
        files: [
          { newUri: vscode.Uri.parse("File1New"), oldUri: vscode.Uri.parse("File1Old") },
          { newUri: vscode.Uri.parse("File2New"), oldUri: vscode.Uri.parse("File2Old") },
        ],
      });
      chai.assert.isTrue(isEnvFileStub.calledOnce);
      chai.assert.isTrue(reloadEnvStub.calledOnce);
    });

    it("No need to refresh Env", async () => {
      mockValue(globalVariables, "core", new MockCore());
      const isEnvFileStub = vi
        .spyOn(globalVariables.core, "isEnvFile")
        .mockResolvedValue(ok(false));
      const reloadEnvStub = vi.spyOn(envTreeProviderInstance, "reloadEnvironments");
      await refreshEnvTreeOnFilesNameChanged("workspaceUri", {
        files: [
          { newUri: vscode.Uri.parse("File1New"), oldUri: vscode.Uri.parse("File1Old") },
          { newUri: vscode.Uri.parse("File2New"), oldUri: vscode.Uri.parse("File2Old") },
        ],
      });
      chai.assert.isTrue(isEnvFileStub.callCount === 4);
      chai.assert.isTrue(reloadEnvStub.notCalled);
    });
  });

  // eslint-disable-next-line no-secrets/no-secrets
  describe("refreshEnvTreeOnProjectSettingFileChanged", function () {
    it("Refresh Env", async () => {
      mockValue(globalVariables, "core", new MockCore());
      const reloadEnvStub = vi.spyOn(envTreeProviderInstance, "reloadEnvironments");
      await refreshEnvTreeOnProjectSettingFileChanged(
        ".",
        path.resolve(".", `.fx`, "configs", "projectSettings.json")
      );
      chai.assert.isTrue(reloadEnvStub.calledOnce);
    });

    it("No need to refresh Env", async () => {
      mockValue(globalVariables, "core", new MockCore());
      const reloadEnvStub = vi.spyOn(envTreeProviderInstance, "reloadEnvironments");
      await refreshEnvTreeOnProjectSettingFileChanged(
        "..",
        path.resolve(".", `.fx`, "configs", "projectSettings.json")
      );
      chai.assert.isTrue(reloadEnvStub.notCalled);
    });
  });
});

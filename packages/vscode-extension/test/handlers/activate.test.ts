import * as sinon from "sinon";
import * as chai from "chai";
import * as vscode from "vscode";
import path from "path";
import * as fileSystemWatcher from "../../src/utils/fileSystemWatcher";
import * as globalVariables from "../../src/globalVariables";
import * as projectSettingsHelper from "@microsoft/teamsfx-core/build/common/projectSettingsHelper";
import {
  activate,
  refreshEnvTreeOnEnvFileChanged,
  refreshEnvTreeOnFilesNameChanged,
  refreshEnvTreeOnProjectSettingFileChanged,
} from "../../src/handlers/activate";
import { ok } from "@microsoft/teamsfx-api";
import { FxCore } from "@microsoft/teamsfx-core";
import commandController from "../../src/commandController";
import { AzureAccountManager } from "../../src/commonlib/azureLogin";
import { signedIn, signedOut } from "../../src/commonlib/common/constant";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import accountTreeViewProviderInstance from "../../src/treeview/account/accountTreeViewProvider";
import envTreeProviderInstance from "../../src//treeview/environmentTreeViewProvider";
import TreeViewManagerInstance from "../../src/treeview/treeViewManager";
import M365TokenInstance from "../../src/commonlib/m365Login";
import { MockCore } from "../mocks/mockCore";

describe("Activate", function () {
  describe("activate()", function () {
    const sandbox = sinon.createSandbox();

    beforeEach(() => {
      sandbox.stub(accountTreeViewProviderInstance, "subscribeToStatusChanges");
      sandbox.stub(vscode.extensions, "getExtension").returns(undefined);
      sandbox.stub(TreeViewManagerInstance, "getTreeView").returns(undefined);
      sandbox.stub(ExtTelemetry, "dispose");
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("No globalState error", async () => {
      const result = await activate();
      chai.assert.deepEqual(result.isOk() ? result.value : result.error.name, {});
    });

    it("Valid project", async () => {
      sandbox.stub(projectSettingsHelper, "isValidProject").returns(true);
      const sendTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      const addSharedPropertyStub = sandbox.stub(ExtTelemetry, "addSharedProperty");
      const setCommandIsRunningStub = sandbox.stub(globalVariables, "setCommandIsRunning");
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.parse("test"));
      const addFileSystemWatcherStub = sandbox.stub(fileSystemWatcher, "addFileSystemWatcher");
      const lockedByOperationStub = sandbox.stub(commandController, "lockedByOperation");
      const unlockedByOperationStub = sandbox.stub(commandController, "unlockedByOperation");
      const azureAccountSetStatusChangeMapStub = sandbox.stub(
        AzureAccountManager.prototype,
        "setStatusChangeMap"
      );
      const m365AccountSetStatusChangeMapStub = sandbox.stub(
        M365TokenInstance,
        "setStatusChangeMap"
      );
      const showMessageStub = sandbox.stub(vscode.window, "showInformationMessage");
      let lockCallback: any;
      let unlockCallback: any;

      sandbox.stub(FxCore.prototype, "on").callsFake((event: string, callback: any) => {
        if (event === "lock") {
          lockCallback = callback;
        } else {
          unlockCallback = callback;
        }
      });
      azureAccountSetStatusChangeMapStub.callsFake(
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
      m365AccountSetStatusChangeMapStub.callsFake(
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

    it("throws error", async () => {
      sandbox.stub(projectSettingsHelper, "isValidProject").returns(false);
      sandbox.stub(M365TokenInstance, "setStatusChangeMap");
      sandbox.stub(FxCore.prototype, "on").throws(new Error("test"));
      const showErrorMessageStub = sandbox.stub(vscode.window, "showErrorMessage");

      const result = await activate();

      chai.assert.isTrue(result.isErr());
      chai.assert.isTrue(showErrorMessageStub.called);
    });
  });

  describe("refreshEnvTreeOnEnvFileChanged", function () {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("Refresh Env", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const isEnvFileStub = sandbox.stub(globalVariables.core, "isEnvFile").resolves(ok(true));
      const reloadEnvStub = sandbox.stub(envTreeProviderInstance, "reloadEnvironments");
      await refreshEnvTreeOnEnvFileChanged("workspaceUri", [
        vscode.Uri.parse("File1"),
        vscode.Uri.parse("File2"),
      ]);
      chai.assert.isTrue(isEnvFileStub.calledOnce);
      chai.assert.isTrue(reloadEnvStub.calledOnce);
    });

    it("No need to refresh Env", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const isEnvFileStub = sandbox.stub(globalVariables.core, "isEnvFile").resolves(ok(false));
      const reloadEnvStub = sandbox.stub(envTreeProviderInstance, "reloadEnvironments");
      await refreshEnvTreeOnEnvFileChanged("workspaceUri", [
        vscode.Uri.parse("File1"),
        vscode.Uri.parse("File2"),
      ]);
      chai.assert.isTrue(isEnvFileStub.calledTwice);
      chai.assert.isTrue(reloadEnvStub.notCalled);
    });
  });

  describe("refreshEnvTreeOnFilesNameChanged", function () {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("Refresh Env", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const isEnvFileStub = sandbox
        .stub(globalVariables.core, "isEnvFile")
        .callsFake((projectPath, inputFile) => {
          if (inputFile === "File1New" || inputFile === "File2New") {
            return Promise.resolve(ok(true));
          }
          return Promise.resolve(ok(false));
        });
      const reloadEnvStub = sandbox.stub(envTreeProviderInstance, "reloadEnvironments");
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
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const isEnvFileStub = sandbox.stub(globalVariables.core, "isEnvFile").resolves(ok(false));
      const reloadEnvStub = sandbox.stub(envTreeProviderInstance, "reloadEnvironments");
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
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("Refresh Env", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const reloadEnvStub = sandbox.stub(envTreeProviderInstance, "reloadEnvironments");
      await refreshEnvTreeOnProjectSettingFileChanged(
        ".",
        path.resolve(".", `.fx`, "configs", "projectSettings.json")
      );
      chai.assert.isTrue(reloadEnvStub.calledOnce);
    });

    it("No need to refresh Env", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const reloadEnvStub = sandbox.stub(envTreeProviderInstance, "reloadEnvironments");
      await refreshEnvTreeOnProjectSettingFileChanged(
        "..",
        path.resolve(".", `.fx`, "configs", "projectSettings.json")
      );
      chai.assert.isTrue(reloadEnvStub.notCalled);
    });
  });
});

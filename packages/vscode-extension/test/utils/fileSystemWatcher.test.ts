import * as sinon from "sinon";
import * as chai from "chai";
import * as globalVariables from "../../src/globalVariables";
import fs from "fs-extra";
import * as vscode from "vscode";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import {
  addFileSystemWatcher,
  fileSystemWatcherDeps,
  refreshSPFxTreeOnFileChanged,
  sendSDKVersionTelemetry,
} from "../../src/utils/fileSystemWatcher";
import TreeViewManagerInstance from "../../src/treeview/treeViewManager";
import * as projectSettingsHelper from "@microsoft/teamsfx-core/build/common/projectSettingsHelper";

describe("FileSystemWatcher", function () {
  describe("addFileSystemWatcher", function () {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("addFileSystemWatcher detect SPFx project", async () => {
      const workspacePath = "test";
      sandbox.stub(fileSystemWatcherDeps, "isValidProject").returns(true);
      sandbox.stub(fileSystemWatcherDeps, "initializeGlobalVariables");
      sandbox.stub(fileSystemWatcherDeps, "updateDevelopmentTreeView");

      const watcher = {
        onDidCreate: () => ({ dispose: () => undefined }),
        onDidChange: () => ({ dispose: () => undefined }),
        onDidDelete: () => ({ dispose: () => undefined }),
      } as any;
      const createWatcher = sandbox
        .stub(fileSystemWatcherDeps, "createFileSystemWatcher")
        .returns(watcher);
      const createListener = sandbox
        .stub(watcher, "onDidCreate")
        .callsFake((...args: unknown[]) => {
          (args as any)[0]({ fsPath: "test/package-lock.json" });
          return { dispose: () => undefined };
        });
      const changeListener = sandbox
        .stub(watcher, "onDidChange")
        .callsFake((...args: unknown[]) => {
          (args as any)[0]({ fsPath: "test/package-lock.json" });
          return { dispose: () => undefined };
        });
      sandbox.stub(watcher, "onDidDelete").callsFake((...args: unknown[]) => {
        (args as any)[0]({ fsPath: "test/.yo-rc.json" });
        return { dispose: () => undefined };
      });
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent").callsFake(() => {});

      addFileSystemWatcher(workspacePath);

      chai.assert.equal(createWatcher.callCount, 2);
      chai.assert.equal(createListener.callCount, 2);
      chai.assert.isTrue(changeListener.calledTwice);
    });

    it("addFileSystemWatcher in invalid project", async () => {
      const workspacePath = "test";
      sandbox.stub(fileSystemWatcherDeps, "isValidProject").returns(false);

      const watcher = {
        onDidCreate: () => ({ dispose: () => undefined }),
        onDidChange: () => ({ dispose: () => undefined }),
      } as any;
      const createWatcher = sandbox
        .stub(fileSystemWatcherDeps, "createFileSystemWatcher")
        .returns(watcher);
      const createListener = sandbox.stub(watcher, "onDidCreate").resolves();
      const changeListener = sandbox.stub(watcher, "onDidChange").resolves();

      addFileSystemWatcher(workspacePath);

      chai.assert.isTrue(createWatcher.notCalled);
      chai.assert.isTrue(createListener.notCalled);
      chai.assert.isTrue(changeListener.notCalled);
    });
  });

  describe("refreshSPFxTreeOnFileChanged", function () {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("refreshSPFxTreeOnFileChanged", () => {
      const initGlobalVariables = sandbox.stub(fileSystemWatcherDeps, "initializeGlobalVariables");
      const updateDevelopmentTreeView = sandbox
        .stub(fileSystemWatcherDeps, "updateDevelopmentTreeView")
        .resolves();

      refreshSPFxTreeOnFileChanged();

      chai.expect(initGlobalVariables.calledOnce).to.be.true;
      chai.expect(updateDevelopmentTreeView.calledOnce).to.be.true;
    });
  });

  describe("sendSDKVersionTelemetry", function () {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("happy path", async () => {
      const filePath = "test/package-lock.json";

      const readJsonFunc = sandbox.stub(fileSystemWatcherDeps, "readJson").resolves();
      sandbox.stub(fileSystemWatcherDeps, "sendTelemetryEvent");

      sendSDKVersionTelemetry(filePath);

      chai.assert.isTrue(readJsonFunc.calledOnce);
    });
  });
});

import * as chai from "chai";
import * as globalVariables from "../../src/globalVariables";
import fs from "fs-extra";
import * as vscode from "vscode";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { vi } from "vitest";
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
    it("addFileSystemWatcher detect SPFx project", async () => {
      const workspacePath = "test";
      vi.spyOn(fileSystemWatcherDeps, "isValidProject").mockReturnValue(true);
      vi.spyOn(fileSystemWatcherDeps, "initializeGlobalVariables").mockImplementation(() => {});
      vi.spyOn(fileSystemWatcherDeps, "updateDevelopmentTreeView").mockResolvedValue();

      const watcher = {
        onDidCreate: () => ({ dispose: () => undefined }),
        onDidChange: () => ({ dispose: () => undefined }),
        onDidDelete: () => ({ dispose: () => undefined }),
      } as any;
      const createWatcher = vi
        .spyOn(fileSystemWatcherDeps, "createFileSystemWatcher")
        .mockReturnValue(watcher);
      const createListener = vi
        .spyOn(watcher, "onDidCreate")
        .mockImplementation((...args: unknown[]) => {
          (args as any)[0]({ fsPath: "test/package-lock.json" });
          return { dispose: () => undefined };
        });
      const changeListener = vi
        .spyOn(watcher, "onDidChange")
        .mockImplementation((...args: unknown[]) => {
          (args as any)[0]({ fsPath: "test/package-lock.json" });
          return { dispose: () => undefined };
        });
      vi.spyOn(watcher, "onDidDelete").mockImplementation((...args: unknown[]) => {
        return { dispose: () => undefined };
      });
      vi.spyOn(ExtTelemetry, "sendTelemetryEvent").mockImplementation(() => {});

      addFileSystemWatcher(workspacePath);

      chai.assert.equal(createWatcher.callCount, 2);
      chai.assert.equal(createListener.callCount, 2);
      chai.assert.isTrue(changeListener.calledTwice);
    });

    it("addFileSystemWatcher in invalid project", async () => {
      const workspacePath = "test";
      vi.spyOn(fileSystemWatcherDeps, "isValidProject").mockReturnValue(false);

      const watcher = {
        onDidCreate: () => ({ dispose: () => undefined }),
        onDidChange: () => ({ dispose: () => undefined }),
      } as any;
      const createWatcher = vi
        .spyOn(fileSystemWatcherDeps, "createFileSystemWatcher")
        .mockReturnValue(watcher);
      const createListener = vi.spyOn(watcher, "onDidCreate").mockResolvedValue();
      const changeListener = vi.spyOn(watcher, "onDidChange").mockResolvedValue();

      addFileSystemWatcher(workspacePath);

      chai.assert.isTrue(createWatcher.notCalled);
      chai.assert.isTrue(createListener.notCalled);
      chai.assert.isTrue(changeListener.notCalled);
    });
  });

  describe("refreshSPFxTreeOnFileChanged", function () {
    it("refreshSPFxTreeOnFileChanged", () => {
      const initGlobalVariables = vi
        .spyOn(fileSystemWatcherDeps, "initializeGlobalVariables")
        .mockImplementation(() => {});
      const updateDevelopmentTreeView = vi
        .spyOn(fileSystemWatcherDeps, "updateDevelopmentTreeView")
        .mockResolvedValue();

      refreshSPFxTreeOnFileChanged();

      chai.expect(initGlobalVariables.calledOnce).to.be.true;
      chai.expect(updateDevelopmentTreeView.calledOnce).to.be.true;
    });
  });

  describe("sendSDKVersionTelemetry", function () {
    it("happy path", async () => {
      const filePath = "test/package-lock.json";

      const readJsonFunc = vi.spyOn(fileSystemWatcherDeps, "readJson").mockResolvedValue();
      vi.spyOn(fileSystemWatcherDeps, "sendTelemetryEvent");

      sendSDKVersionTelemetry(filePath);

      chai.assert.isTrue(readJsonFunc.calledOnce);
    });
  });
});

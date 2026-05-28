import * as chai from "chai";
import fs from "fs-extra";
import os from "os";
import * as sinon from "sinon";
import cp from "child_process";
import * as vscode from "vscode";
import * as globalVariables from "../../src/globalVariables";
import {
  openFolderInExplorer,
  isWindows,
  isLinux,
  isMacOS,
  hasAdaptiveCardInWorkspace,
  acpInstalled,
  getLocalDebugMessageTemplate,
} from "../../src/utils/commonUtils";
import mockfs from "mock-fs";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";

describe("CommonUtils", () => {
  afterEach(() => {
    // Restore the default sandbox here
    sinon.restore();
  });

  describe("openFolderInExplorer", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("happy path", () => {
      const folderPath = "fakePath";
      sandbox.stub(cp, "exec");
      openFolderInExplorer(folderPath);
    });
  });

  describe("os assertion", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("should return exactly result according to os.type", async () => {
      sandbox.stub(os, "type").returns("Windows_NT");
      chai.expect(isWindows()).equals(true);
      sandbox.restore();

      sandbox.stub(os, "type").returns("Linux");
      chai.expect(isLinux()).equals(true);
      sandbox.restore();

      sandbox.stub(os, "type").returns("Darwin");
      chai.expect(isMacOS()).equals(true);
      sandbox.restore();
    });
  });

  describe("hasAdaptiveCardInWorkspace()", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      mockfs.restore();
      sandbox.restore();
    });

    it("no workspace", async () => {
      sandbox.stub(globalVariables, "workspaceUri").value(undefined);

      const result = await hasAdaptiveCardInWorkspace();

      chai.assert.isFalse(result);
    });

    it("happy path", async () => {
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("/test"));
      mockfs({
        "/test/card.json": JSON.stringify({
          $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
          type: "AdaptiveCard",
          version: "1.5",
          actions: [
            {
              type: "Action.OpenUrl",
              title: "More Info",
              url: "https://example.com",
            },
          ],
        }),
      });

      const result = await hasAdaptiveCardInWorkspace();

      chai.assert.isTrue(result);
    });

    it("hasAdaptiveCardInWorkspace() no adaptive card file", async () => {
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("/test"));
      mockfs({
        "/test/card.json": JSON.stringify({ hello: "world" }),
      });

      const result = await hasAdaptiveCardInWorkspace();

      chai.assert.isFalse(result);
    });

    it("hasAdaptiveCardInWorkspace() very large adaptive card file", async () => {
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("/test"));
      mockfs({
        "/test/card.json": JSON.stringify({
          $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
          type: "AdaptiveCard",
          version: "1.5",
          actions: [
            {
              type: "Action.OpenUrl",
              title: "a".repeat(1024 * 1024 + 10),
              url: "https://example.com",
            },
          ],
        }),
      });

      const result = await hasAdaptiveCardInWorkspace();

      chai.assert.isFalse(result);
    });
  });

  describe("acpInstalled()", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      mockfs.restore();
      sandbox.restore();
    });

    it("already installed", async () => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(vscode.extensions, "getExtension").returns({} as any);

      const installed = acpInstalled();

      chai.assert.isTrue(installed);
    });

    it("not installed", async () => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(vscode.extensions, "getExtension").returns(undefined);

      const installed = acpInstalled();

      chai.assert.isFalse(installed);
    });
  });

  describe("getLocalDebugMessageTemplate()", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("Test Tool enabled in Windows platform", async () => {
      sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
      sandbox.stub(fs, "pathExists").resolves(true);

      const result = await getLocalDebugMessageTemplate(true);
      chai.assert.isTrue(result.includes("Test Tool"));
    });

    it("Test Tool disabled in Windows platform", async () => {
      sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
      sandbox.stub(fs, "pathExists").resolves(false);

      const result = await getLocalDebugMessageTemplate(true);
      chai.assert.isFalse(result.includes("Test Tool"));
    });

    it("Test Tool enabled in non-Windows platform", async () => {
      sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
      sandbox.stub(fs, "pathExists").resolves(true);

      const result = await getLocalDebugMessageTemplate(false);
      chai.assert.isTrue(result.includes("Test Tool"));
    });

    it("Test Tool disabled in non-Windows platform", async () => {
      sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
      sandbox.stub(fs, "pathExists").resolves(false);

      const result = await getLocalDebugMessageTemplate(false);
      chai.assert.isFalse(result.includes("Test Tool"));
    });

    it("No workspace folder", async () => {
      sandbox.stub(vscode.workspace, "workspaceFolders").value([]);
      sandbox.stub(fs, "pathExists").resolves(false);

      const result = await getLocalDebugMessageTemplate(false);
      chai.assert.isFalse(result.includes("Test Tool"));
    });
  });
});

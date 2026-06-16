// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { vi, describe, it, beforeEach, expect } from "vitest";
import * as vscode from "vscode";
import fs from "fs-extra";
import { setupMCPServer } from "../../src/utils/mcpUtils";
import { TelemetryEvent, TelemetryProperty } from "../../src/telemetry/extTelemetryEvents";
import { fsAdapter, pathAdapter } from "../../src/common/npmPackageDeps";
import * as globalVariables from "../../src/globalVariables";
import * as localizeUtils from "../../src/utils/localizeUtils";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { mockValue } from "../mocks/vitestMockUtils";

vi.mock("vscode");

describe("mcpUtils", () => {
  describe("setupMCPServer", () => {
    beforeEach(() => {
      vi.clearAllMocks();
      mockValue(globalVariables, "workspaceUri", vscode.Uri.file("/test"));
      mockValue(globalVariables, "context", { extensionPath: "/mock/extension/path" });

      vi.spyOn(localizeUtils, "localize").mockImplementation((key: string) => {
        const localizationMap: Record<string, string> = {
          "teamstoolkit.mcpUtils.setupMcpServer.message": "Setup MCP Server for %s",
          "teamstoolkit.mcpUtils.setupMcpServer.confirm": "Confirm",
          "teamstoolkit.mcpUtils.setupMcpServer.skip": "Skip",
          "teamstoolkit.mcpUtils.setupMcpServer.successMessage":
            "MCP Server setup completed successfully!",
          "teamstoolkit.mcpUtils.setupMcpServer.errorMessage": "Error setting up MCP Server: %s",
        };
        return localizationMap[key] || key;
      });
    });

    it("should return early if no workspace folders", async () => {
      mockValue(globalVariables, "workspaceUri", undefined);
      const showInformationMessageStub = vi.spyOn(vscode.window, "showInformationMessage");
      const sendTelemetryEventStub = vi.spyOn(ExtTelemetry, "sendTelemetryEvent");

      await setupMCPServer();

      expect(showInformationMessageStub.called).to.be.false;
      expect(sendTelemetryEventStub.called).to.be.false;
    });

    it("should create files when user confirms and both are missing", async () => {
      vi.spyOn(fsAdapter, "existsSync").mockReturnValue(false);
      vi.spyOn(fsAdapter, "mkdirSync").mockImplementation(() => undefined);
      vi.spyOn(fsAdapter, "readFileSync").mockReturnValue("Default copilot instructions content");
      vi.spyOn(fsAdapter, "writeFileSync").mockImplementation(() => undefined);

      const getConfigurationStub = vi.spyOn(vscode.workspace, "getConfiguration");
      const mockMcpConfig = {
        get: vi.fn((key: string) =>
          key === "servers"
            ? {
                "other-server": { command: "other-command" },
              }
            : undefined
        ),
      };
      getConfigurationStub.mockReturnValue(mockMcpConfig as any);

      const showInformationMessageStub = vi
        .spyOn(vscode.window, "showInformationMessage")
        .mockImplementation(async (_message: string, ...items: string[]) => items[0]);

      const sendTelemetryEventStub = vi.spyOn(ExtTelemetry, "sendTelemetryEvent");

      await setupMCPServer();

      expect(sendTelemetryEventStub.called).to.be.true;
    });

    it("should handle error during file creation", async () => {
      vi.spyOn(fsAdapter, "existsSync").mockReturnValue(false);
      vi.spyOn(fsAdapter, "readFileSync").mockImplementation(() => {
        throw new Error("Test error");
      });

      const getConfigurationStub = vi.spyOn(vscode.workspace, "getConfiguration");
      const mockMcpConfig = {
        get: vi.fn((key: string) =>
          key === "servers"
            ? {
                "other-server": { command: "other-command" },
              }
            : undefined
        ),
      };
      getConfigurationStub.mockReturnValue(mockMcpConfig as any);

      const showInformationMessageStub = vi
        .spyOn(vscode.window, "showInformationMessage")
        .mockResolvedValue("Confirm" as any);
      const showErrorMessageStub = vi
        .spyOn(vscode.window, "showErrorMessage")
        .mockResolvedValue(undefined as any);

      const sendTelemetryErrorEventStub = vi.spyOn(ExtTelemetry, "sendTelemetryErrorEvent");

      await setupMCPServer();

      expect(showErrorMessageStub.called).to.be.true;
      expect(sendTelemetryErrorEventStub.called).to.be.true;
    });

    it("should prompt user when copilot instructions file is missing", async () => {
      vi.spyOn(fsAdapter, "existsSync").mockImplementation((filePath: string) => {
        return !filePath.endsWith("copilot-instructions.md");
      });

      const getConfigurationStub = vi.spyOn(vscode.workspace, "getConfiguration");
      const mockMcpConfig = {
        get: vi.fn((key: string) =>
          key === "servers"
            ? {
                "test-server": {
                  command: "npx",
                  args: ["@microsoft/m365agentstoolkit-mcp@latest", "server", "start"],
                },
              }
            : undefined
        ),
      };
      getConfigurationStub.mockReturnValue(mockMcpConfig as any);

      const showInformationMessageStub = vi
        .spyOn(vscode.window, "showInformationMessage")
        .mockResolvedValue("Skip" as any);

      const sendTelemetryEventStub = vi.spyOn(ExtTelemetry, "sendTelemetryEvent");

      await setupMCPServer();

      expect(showInformationMessageStub.called).to.be.true;
      expect(sendTelemetryEventStub.called).to.be.true;
    });
  });

  describe("Adapter delegation tests", () => {
    it("fsAdapter.existsSync delegates to fs-extra", () => {
      const result = fsAdapter.existsSync("/test/path");
      expect(typeof result).to.equal("boolean");
    });

    it("fsAdapter.writeFileSync delegates to fs-extra", () => {
      const writeSpy = vi.spyOn(fs, "writeFileSync").mockImplementation(() => undefined);
      fsAdapter.writeFileSync("/test/path", "content");
      expect(writeSpy.mock.calls.length).to.equal(1);
      expect(writeSpy.mock.calls[0]).to.deep.equal(["/test/path", "content", "utf-8"]);
      writeSpy.mockRestore();
    });

    it("pathAdapter.join delegates to path module", () => {
      const result = pathAdapter.join("/test", "path");
      expect(typeof result).to.equal("string");
      expect(result.toLowerCase()).to.include("test");
    });
  });
});

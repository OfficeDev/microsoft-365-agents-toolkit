import { vi } from "vitest";
import { mockValue } from "../mocks/vitestMockUtils";
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as chai from "chai";
import * as vscode from "vscode";
import fs from "fs-extra";
import * as globalVariables from "../../src/globalVariables";
import { mcpUtilsDeps, setupMCPServer } from "../../src/utils/mcpUtils";
import { TelemetryEvent, TelemetryProperty } from "../../src/telemetry/extTelemetryEvents";
import path from "path";
import { FxError } from "@microsoft/teamsfx-api";

describe("mcpUtils", () => {
  describe("setupMCPServer", () => {
    let showInformationMessageStub: ReturnType<typeof vi.spyOn>;
    let sendTelemetryEventStub: ReturnType<typeof vi.spyOn>;
    let getConfigurationStub: ReturnType<typeof vi.spyOn>;
    let existsSyncStub: ReturnType<typeof vi.spyOn>;
    let localizeStub: ReturnType<typeof vi.spyOn>;

    beforeEach(() => {
      mockValue(globalVariables, "workspaceUri", vscode.Uri.file("/test"));
      mockValue(globalVariables, "context", { extensionPath: "/mock/extension/path" });
      showInformationMessageStub = vi
        .spyOn(mcpUtilsDeps, "showInformationMessage")
        .mockResolvedValue(undefined);
      sendTelemetryEventStub = vi
        .spyOn(mcpUtilsDeps, "sendTelemetryEvent")
        .mockImplementation(() => {
          return;
        });
      getConfigurationStub = vi.spyOn(vscode.workspace, "getConfiguration");
      vi.spyOn(mcpUtilsDeps, "getMcpServers").mockImplementation(() =>
        (vscode.workspace.getConfiguration("mcp") as any).get("servers")
      );
      existsSyncStub = vi.spyOn(mcpUtilsDeps, "existsSync");
      localizeStub = vi.spyOn(mcpUtilsDeps, "localize");
      localizeStub.mockImplementation((key: string) => {
        switch (key) {
          case "teamstoolkit.mcpUtils.setupMcpServer.message":
            return "Setup MCP Server for %s";
          case "teamstoolkit.mcpUtils.setupMcpServer.confirm":
            return "Confirm";
          case "teamstoolkit.mcpUtils.setupMcpServer.skip":
            return "Skip";
          case "teamstoolkit.mcpUtils.setupMcpServer.successMessage":
            return "MCP Server setup completed successfully!";
          case "teamstoolkit.mcpUtils.setupMcpServer.errorMessage":
            return "Error setting up MCP Server: %s";
          default:
            return key;
        }
      });
    });

    it("should return early if no workspace folders", async () => {
      mockValue(globalVariables, "workspaceUri", undefined);

      await setupMCPServer();

      chai.expect(showInformationMessageStub.called).to.be.false;
      chai.expect(sendTelemetryEventStub.called).to.be.false;
    });

    it("should return early if both files exist and MCP config is up to date", async () => {
      existsSyncStub.mockReturnValue(true);

      const mockMcpConfig = {
        get: vi
          .spyOn()
          .withArgs("servers")
          .mockReturnValue({
            "test-server": {
              command: "npx",
              args: ["@microsoft/m365agentstoolkit-mcp@latest", "server", "start"],
            },
          }),
      };
      getConfigurationStub.withArgs("mcp").mockReturnValue(mockMcpConfig);

      await setupMCPServer();

      chai.expect(showInformationMessageStub.called).to.be.false;
      chai.expect(sendTelemetryEventStub.called).to.be.false;
    });

    it("should prompt user when copilot instructions file is missing", async () => {
      existsSyncStub.mockImplementation((filePath: string) => {
        return filePath.endsWith("copilot-instructions.md") ? false : true;
      });
      const mockMcpConfig = {
        get: vi
          .spyOn()
          .withArgs("servers")
          .mockReturnValue({
            "test-server": {
              command: "npx",
              args: ["@microsoft/m365agentstoolkit-mcp@latest", "server", "start"],
            },
          }),
      };
      getConfigurationStub.withArgs("mcp").mockReturnValue(mockMcpConfig);

      showInformationMessageStub.mockResolvedValue("Skip");

      await setupMCPServer();

      chai.expect(
        sendTelemetryEventStub.calledWith(TelemetryEvent.PromptMCPServer, {
          [TelemetryProperty.MissingCopilotInstructions]: "true",
          [TelemetryProperty.MissingMCPConfig]: "false",
        })
      ).to.be.true;

      chai.expect(showInformationMessageStub.called).to.be.true;

      chai.expect(
        sendTelemetryEventStub.calledWith(TelemetryEvent.PromptMCPServer, {
          [TelemetryProperty.UserSelection]: "skip",
        })
      ).to.be.true;
    });

    it("should prompt user when MCP config needs update", async () => {
      existsSyncStub.mockImplementation((filePath: string) => {
        return filePath.endsWith("copilot-instructions.md") ? true : false;
      });

      (mcpUtilsDeps.getMcpServers as any).mockReturnValue({
        "other-server": {
          command: "other-command",
        },
      });

      showInformationMessageStub.mockResolvedValue("Skip");

      await setupMCPServer();

      chai.expect(
        sendTelemetryEventStub.calledWith(TelemetryEvent.PromptMCPServer, {
          [TelemetryProperty.MissingCopilotInstructions]: "false",
          [TelemetryProperty.MissingMCPConfig]: "true",
        })
      ).to.be.true;

      chai.expect(showInformationMessageStub.called).to.be.true;

      chai.expect(
        sendTelemetryEventStub.calledWith(TelemetryEvent.PromptMCPServer, {
          [TelemetryProperty.UserSelection]: "skip",
        })
      ).to.be.true;
    });

    it("should prompt user when both files are missing", async () => {
      existsSyncStub.mockReturnValue(false);

      const mockMcpConfig = {
        get: vi
          .spyOn()
          .withArgs("servers")
          .mockReturnValue({
            "other-server": {
              command: "other-command",
            },
          }),
      };
      getConfigurationStub.withArgs("mcp").mockReturnValue(mockMcpConfig);

      showInformationMessageStub.mockResolvedValue("Skip");

      await setupMCPServer();

      chai.expect(
        sendTelemetryEventStub.calledWith(TelemetryEvent.PromptMCPServer, {
          [TelemetryProperty.MissingCopilotInstructions]: "true",
          [TelemetryProperty.MissingMCPConfig]: "true",
        })
      ).to.be.true;

      chai.expect(showInformationMessageStub.called).to.be.true;
    });

    it("should create files when user confirms and both are missing", async () => {
      existsSyncStub.mockReturnValue(false);
      const mkdirSyncStub = vi.spyOn(mcpUtilsDeps, "mkdirSync").mockImplementation(() => undefined);
      const readFileSyncStub = vi.spyOn(mcpUtilsDeps, "readFileSync");
      const writeFileSyncStub = vi
        .spyOn(mcpUtilsDeps, "writeFileSync")
        .mockImplementation(() => {});
      const openSyncStub = vi.spyOn(mcpUtilsDeps, "openSync");
      vi.spyOn(mcpUtilsDeps, "closeSync").mockImplementation(() => {});
      const showErrorMessageStub = vi
        .spyOn(mcpUtilsDeps, "showErrorMessage")
        .mockResolvedValue(undefined);

      readFileSyncStub.mockReturnValue("Default copilot instructions content");
      openSyncStub.mockReturnValue(3);

      const mockMcpConfig = {
        get: vi.fn((key: string) =>
          key === "servers"
            ? {
                "other-server": {
                  command: "other-command",
                },
              }
            : undefined
        ),
      };
      getConfigurationStub.mockReturnValue(mockMcpConfig as any);

      showInformationMessageStub.mockImplementation(
        async (_message: string, ...items: string[]) => {
          return items[0];
        }
      );

      await setupMCPServer();

      chai.expect(
        mkdirSyncStub.mock.calls.some(
          (call) => typeof call[0] === "string" && (call[0] as string).endsWith(".github")
        )
      ).to.be.true;
      chai.expect(writeFileSyncStub.called).to.be.true;
    });

    it("should create only copilot instructions file when MCP config is up to date", async () => {
      existsSyncStub.mockImplementation((filePath: string) => {
        if (filePath.endsWith("copilot-instructions.md")) return false;
        if (filePath.endsWith(".github")) return false;
        return true;
      });

      const mkdirSyncStub = vi.spyOn(mcpUtilsDeps, "mkdirSync").mockImplementation(() => undefined);
      const readFileSyncStub = vi.spyOn(mcpUtilsDeps, "readFileSync");
      const writeFileSyncStub = vi
        .spyOn(mcpUtilsDeps, "writeFileSync")
        .mockImplementation(() => {});
      const openSyncStub = vi.spyOn(mcpUtilsDeps, "openSync");
      vi.spyOn(mcpUtilsDeps, "closeSync").mockImplementation(() => {});

      readFileSyncStub.mockReturnValue("Default copilot instructions content");
      openSyncStub.mockReturnValue(3);

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

      showInformationMessageStub.mockImplementation(
        async (_message: string, ...items: string[]) => {
          return items[0];
        }
      );

      await setupMCPServer();

      chai.expect(mkdirSyncStub.calledOnce).to.be.true;
      chai.expect(writeFileSyncStub.calledOnce).to.be.true;
      chai.expect(readFileSyncStub.calledOnce).to.be.true;
    });

    it("should create only MCP config file when copilot instructions exist", async () => {
      existsSyncStub.mockImplementation((filePath: string) => {
        if (filePath.endsWith("copilot-instructions.md")) return true;
        if (filePath.endsWith("mcp.json")) return false;
        if (filePath.endsWith(".vscode")) return false;
        return true;
      });

      const mkdirSyncStub = vi.spyOn(mcpUtilsDeps, "mkdirSync").mockImplementation(() => undefined);
      const writeFileSyncStub = vi
        .spyOn(mcpUtilsDeps, "writeFileSync")
        .mockImplementation(() => {});

      const mockMcpConfig = {
        get: vi.fn((key: string) =>
          key === "servers"
            ? {
                "other-server": {
                  command: "other-command",
                },
              }
            : undefined
        ),
      };
      getConfigurationStub.mockReturnValue(mockMcpConfig as any);

      showInformationMessageStub.mockImplementation(
        async (_message: string, ...items: string[]) => {
          return items[0];
        }
      );

      await setupMCPServer();

      chai.expect(mkdirSyncStub.calledOnce).to.be.true;
      chai.expect(writeFileSyncStub.calledOnce).to.be.true;

      const writtenConfig = JSON.parse(writeFileSyncStub.firstCall.args[1] as string);
      chai.expect(writtenConfig.servers.m365agentstoolkit).to.deep.equal({
        command: "npx",
        args: ["@microsoft/m365agentstoolkit-mcp@latest", "server", "start"],
      });
    });

    it("should update existing MCP config file when it exists", async () => {
      existsSyncStub.mockImplementation((filePath: string) => {
        if (filePath.endsWith("copilot-instructions.md")) return true;
        if (filePath.endsWith("mcp.json")) return true;
        return true;
      });

      const readFileSyncStub = vi.spyOn(mcpUtilsDeps, "readFileSync");
      const writeFileSyncStub = vi.spyOn(mcpUtilsDeps, "writeFileSync");
      const openSyncStub = vi.spyOn(mcpUtilsDeps, "openSync");
      vi.spyOn(mcpUtilsDeps, "closeSync");

      readFileSyncStub.mockReturnValue(
        JSON.stringify({
          servers: {
            "existing-server": {
              command: "existing-command",
            },
          },
        })
      );
      openSyncStub.mockReturnValue(3);

      const mockMcpConfig = {
        get: vi
          .spyOn()
          .withArgs("servers")
          .mockReturnValue({
            "other-server": {
              command: "other-command",
            },
          }),
      };
      getConfigurationStub.withArgs("mcp").mockReturnValue(mockMcpConfig);

      showInformationMessageStub.mockResolvedValue("Confirm");

      await setupMCPServer();

      chai.expect(readFileSyncStub.calledOnce).to.be.true;
      chai.expect(writeFileSyncStub.calledOnce).to.be.true;
      chai.expect(openSyncStub.calledOnce).to.be.true;

      const writtenConfig = JSON.parse(writeFileSyncStub.firstCall.args[1] as string);
      chai.expect(writtenConfig.servers["M365AgentsToolkit MCP Server"]).to.deep.equal({
        command: "npx",
        args: ["@microsoft/m365agentstoolkit-mcp@latest", "server", "start"],
      });
    });

    it("should handle MCP config file with no servers object", async () => {
      existsSyncStub.mockImplementation((filePath: string) => {
        if (filePath.endsWith("copilot-instructions.md")) return true;
        if (filePath.endsWith("mcp.json")) return true;
        return true;
      });

      const readFileSyncStub = vi.spyOn(mcpUtilsDeps, "readFileSync");
      const writeFileSyncStub = vi.spyOn(mcpUtilsDeps, "writeFileSync");
      const openSyncStub = vi.spyOn(mcpUtilsDeps, "openSync");
      vi.spyOn(mcpUtilsDeps, "closeSync");

      readFileSyncStub.mockReturnValue(
        JSON.stringify({
          someOtherProperty: "value",
        })
      );
      openSyncStub.mockReturnValue(3);

      const mockMcpConfig = {
        get: vi
          .spyOn()
          .withArgs("servers")
          .mockReturnValue({
            "other-server": {
              command: "other-command",
            },
          }),
      };
      getConfigurationStub.withArgs("mcp").mockReturnValue(mockMcpConfig);

      showInformationMessageStub.mockResolvedValue("Confirm");

      await setupMCPServer();

      const writtenConfig = JSON.parse(writeFileSyncStub.firstCall.args[1] as string);
      chai.expect(writtenConfig.servers).to.exist;
      chai.expect(writtenConfig.servers["M365AgentsToolkit MCP Server"]).to.deep.equal({
        command: "npx",
        args: ["@microsoft/m365agentstoolkit-mcp@latest", "server", "start"],
      });
    });

    it("should handle MCP config file with null servers object", async () => {
      existsSyncStub.mockImplementation((filePath: string) => {
        if (filePath.endsWith("copilot-instructions.md")) return true;
        if (filePath.endsWith("mcp.json")) return true;
        return true;
      });

      const readFileSyncStub = vi.spyOn(mcpUtilsDeps, "readFileSync");
      const writeFileSyncStub = vi.spyOn(mcpUtilsDeps, "writeFileSync");
      const openSyncStub = vi.spyOn(mcpUtilsDeps, "openSync");
      vi.spyOn(mcpUtilsDeps, "closeSync");

      readFileSyncStub.mockReturnValue(
        JSON.stringify({
          servers: null,
        })
      );
      openSyncStub.mockReturnValue(3);

      const mockMcpConfig = {
        get: vi
          .spyOn()
          .withArgs("servers")
          .mockReturnValue({
            "other-server": {
              command: "other-command",
            },
          }),
      };
      getConfigurationStub.withArgs("mcp").mockReturnValue(mockMcpConfig);

      showInformationMessageStub.mockResolvedValue("Confirm");

      await setupMCPServer();

      const writtenConfig = JSON.parse(writeFileSyncStub.firstCall.args[1] as string);
      chai.expect(writtenConfig.servers).to.exist;
      chai.expect(writtenConfig.servers["M365AgentsToolkit MCP Server"]).to.deep.equal({
        command: "npx",
        args: ["@microsoft/m365agentstoolkit-mcp@latest", "server", "start"],
      });
    });

    it("should handle error during file creation", async () => {
      existsSyncStub.mockReturnValue(false);
      vi.spyOn(mcpUtilsDeps, "mkdirSync");
      const readFileSyncStub = vi.spyOn(mcpUtilsDeps, "readFileSync");
      const showErrorMessageStub = vi
        .spyOn(mcpUtilsDeps, "showErrorMessage")
        .mockResolvedValue(undefined);
      const sendTelemetryErrorEventStub = vi.spyOn(mcpUtilsDeps, "sendTelemetryErrorEvent");
      sendTelemetryErrorEventStub.mockImplementation(() => {
        return;
      });

      const testError = new Error("Test error");
      readFileSyncStub.throws(testError);

      const mockMcpConfig = {
        get: vi
          .spyOn()
          .withArgs("servers")
          .mockReturnValue({
            "other-server": {
              command: "other-command",
            },
          }),
      };
      getConfigurationStub.withArgs("mcp").mockReturnValue(mockMcpConfig);

      showInformationMessageStub.mockResolvedValue("Confirm");

      await setupMCPServer();

      chai.expect(showErrorMessageStub.called).to.be.true;
      chai.expect(
        sendTelemetryErrorEventStub.calledWith(
          TelemetryEvent.PromptMCPServer,
          testError as FxError,
          {
            [TelemetryProperty.UserSelection]: "confirm",
          }
        )
      ).to.be.true;
    });

    it("should handle error when getConfiguration throws", async () => {
      existsSyncStub.mockImplementation((filePath: string) => {
        return filePath.endsWith("copilot-instructions.md") ? true : false;
      });

      getConfigurationStub.withArgs("mcp").throws(new Error("Config error"));

      showInformationMessageStub.mockResolvedValue("Skip");

      await setupMCPServer();

      chai.expect(
        sendTelemetryEventStub.calledWith(TelemetryEvent.PromptMCPServer, {
          [TelemetryProperty.MissingCopilotInstructions]: "false",
          [TelemetryProperty.MissingMCPConfig]: "true",
        })
      ).to.be.true;
    });

    it("should handle null servers configuration", async () => {
      existsSyncStub.mockImplementation((filePath: string) => {
        return filePath.endsWith("copilot-instructions.md") ? true : false;
      });

      const mockMcpConfig = {
        get: vi.fn().withArgs("servers").mockReturnValue(null),
      };
      getConfigurationStub.withArgs("mcp").mockReturnValue(mockMcpConfig);

      showInformationMessageStub.mockResolvedValue("Skip");

      await setupMCPServer();

      chai.expect(
        sendTelemetryEventStub.calledWith(TelemetryEvent.PromptMCPServer, {
          [TelemetryProperty.MissingCopilotInstructions]: "false",
          [TelemetryProperty.MissingMCPConfig]: "true",
        })
      ).to.be.true;
    });

    it("should handle undefined servers configuration", async () => {
      existsSyncStub.mockImplementation((filePath: string) => {
        return filePath.endsWith("copilot-instructions.md") ? true : false;
      });

      const mockMcpConfig = {
        get: vi.fn().withArgs("servers").mockReturnValue(undefined),
      };
      getConfigurationStub.withArgs("mcp").mockReturnValue(mockMcpConfig);

      showInformationMessageStub.mockResolvedValue("Skip");

      await setupMCPServer();

      chai.expect(
        sendTelemetryEventStub.calledWith(TelemetryEvent.PromptMCPServer, {
          [TelemetryProperty.MissingCopilotInstructions]: "false",
          [TelemetryProperty.MissingMCPConfig]: "true",
        })
      ).to.be.true;
    });

    it("should create github directory when it doesn't exist", async () => {
      existsSyncStub.mockImplementation((filePath: string) => {
        if (filePath.endsWith("copilot-instructions.md")) return false;
        if (filePath.endsWith(".github")) return false;
        return true;
      });

      const mkdirSyncStub = vi.spyOn(mcpUtilsDeps, "mkdirSync").mockImplementation(() => undefined);
      const readFileSyncStub = vi.spyOn(mcpUtilsDeps, "readFileSync");
      const writeFileSyncStub = vi
        .spyOn(mcpUtilsDeps, "writeFileSync")
        .mockImplementation(() => {});
      const openSyncStub = vi.spyOn(mcpUtilsDeps, "openSync");
      vi.spyOn(mcpUtilsDeps, "closeSync").mockImplementation(() => {});

      readFileSyncStub.mockReturnValue("Default copilot instructions content");
      openSyncStub.mockReturnValue(3);

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

      showInformationMessageStub.mockImplementation(
        async (_message: string, ...items: string[]) => {
          return items[0];
        }
      );

      await setupMCPServer();

      chai.expect(mkdirSyncStub.calledWith(path.join("/test", ".github"), { recursive: true })).to
        .be.true;
      chai.expect(writeFileSyncStub.calledOnce).to.be.true;
      chai.expect(readFileSyncStub.calledOnce).to.be.true;
      chai.expect(openSyncStub.calledOnce).to.be.true;
    });

    it("should create vscode directory when it doesn't exist", async () => {
      existsSyncStub.mockImplementation((filePath: string) => {
        if (filePath.endsWith("copilot-instructions.md")) return true;
        if (filePath.endsWith("mcp.json")) return false;
        if (filePath.endsWith(".vscode")) return false;
        return true;
      });

      const mkdirSyncStub = vi.spyOn(mcpUtilsDeps, "mkdirSync").mockImplementation(() => undefined);
      const writeFileSyncStub = vi
        .spyOn(mcpUtilsDeps, "writeFileSync")
        .mockImplementation(() => {});

      const mockMcpConfig = {
        get: vi.fn((key: string) =>
          key === "servers"
            ? {
                "other-server": {
                  command: "other-command",
                },
              }
            : undefined
        ),
      };
      getConfigurationStub.mockReturnValue(mockMcpConfig as any);

      showInformationMessageStub.mockImplementation(
        async (_message: string, ...items: string[]) => {
          return items[0];
        }
      );

      await setupMCPServer();

      chai.expect(mkdirSyncStub.calledWith(path.join("/test", ".vscode"), { recursive: true })).to
        .be.true;
      chai.expect(writeFileSyncStub.calledOnce).to.be.true;

      const writtenConfig = JSON.parse(writeFileSyncStub.firstCall.args[1] as string);
      chai.expect(writtenConfig.servers.m365agentstoolkit).to.deep.equal({
        command: "npx",
        args: ["@microsoft/m365agentstoolkit-mcp@latest", "server", "start"],
      });
    });

    it("should format message correctly for two files", async () => {
      existsSyncStub.mockReturnValue(false);

      const mockMcpConfig = {
        get: vi
          .spyOn()
          .withArgs("servers")
          .mockReturnValue({
            "other-server": {
              command: "other-command",
            },
          }),
      };
      getConfigurationStub.withArgs("mcp").mockReturnValue(mockMcpConfig);

      showInformationMessageStub.mockResolvedValue("Skip");

      await setupMCPServer();

      chai.expect(showInformationMessageStub.called).to.be.true;
      const messageCall = showInformationMessageStub.firstCall.args[0];
      chai.expect(messageCall).to.include("and");
    });
  });

  describe("mcpUtilsDeps", () => {
    it("existsSync delegates to fs", () => {
      vi.spyOn(fs, "existsSync").mockReturnValue(true);
      const result = mcpUtilsDeps.existsSync("/test");
      chai.expect(result).to.be.true;
    });

    it("mkdirSync delegates to fs", () => {
      const spy = vi.spyOn(fs, "mkdirSync").mockImplementation((() => undefined) as any);
      mcpUtilsDeps.mkdirSync("/test", { recursive: true });
      chai.expect(spy.called).to.be.true;
    });

    it("readFileSync delegates to fs", () => {
      vi.spyOn(fs, "readFileSync").mockReturnValue("content" as any);
      const result = mcpUtilsDeps.readFileSync("/test", "utf8");
      chai.expect(result).to.equal("content");
    });

    it("writeFileSync delegates to fs", () => {
      const spy = vi.spyOn(fs, "writeFileSync").mockImplementation(() => {});
      mcpUtilsDeps.writeFileSync("/test", "data", "utf8");
      chai.expect(spy.called).to.be.true;
    });

    it("openSync delegates to fs", () => {
      vi.spyOn(fs, "openSync").mockReturnValue(5 as any);
      const result = mcpUtilsDeps.openSync("/test", "r");
      chai.expect(result).to.equal(5);
    });

    it("closeSync delegates to fs", () => {
      const spy = vi.spyOn(fs, "closeSync").mockImplementation(() => {});
      mcpUtilsDeps.closeSync(5);
      chai.expect(spy.called).to.be.true;
    });

    it("getMcpServers delegates to vscode configuration", () => {
      vi.spyOn(vscode.workspace, "getConfiguration").mockReturnValue({
        get: vi.fn().mockReturnValue({ "test-server": {} }),
      } as any);
      const result = mcpUtilsDeps.getMcpServers();
      chai.expect(result).to.deep.equal({ "test-server": {} });
    });
  });
});

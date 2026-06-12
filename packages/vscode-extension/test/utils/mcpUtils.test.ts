// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";
import * as globalVariables from "../../src/globalVariables";
import { mcpUtilsDeps, setupMCPServer } from "../../src/utils/mcpUtils";
import { TelemetryEvent, TelemetryProperty } from "../../src/telemetry/extTelemetryEvents";
import path from "path";
import { FxError } from "@microsoft/teamsfx-api";

describe("mcpUtils", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sinon.restore();
  });

  describe("setupMCPServer", () => {
    let showInformationMessageStub: sinon.SinonStub;
    let sendTelemetryEventStub: sinon.SinonStub;
    let getConfigurationStub: sinon.SinonStub;
    let existsSyncStub: sinon.SinonStub;
    let localizeStub: sinon.SinonStub;

    beforeEach(() => {
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("/test"));
      sandbox.stub(globalVariables, "context").value({ extensionPath: "/mock/extension/path" });
      showInformationMessageStub = sandbox
        .stub(mcpUtilsDeps, "showInformationMessage")
        .resolves(undefined);
      sendTelemetryEventStub = sandbox.stub(mcpUtilsDeps, "sendTelemetryEvent");
      getConfigurationStub = sandbox.stub(vscode.workspace, "getConfiguration");
      sandbox
        .stub(mcpUtilsDeps, "getMcpServers")
        .callsFake(() => (vscode.workspace.getConfiguration("mcp") as any).get("servers"));
      existsSyncStub = sandbox.stub(mcpUtilsDeps, "existsSync");
      localizeStub = sandbox.stub(mcpUtilsDeps, "localize");
      localizeStub
        .withArgs("teamstoolkit.mcpUtils.setupMcpServer.message")
        .returns("Setup MCP Server for %s");
      localizeStub.withArgs("teamstoolkit.mcpUtils.setupMcpServer.confirm").returns("Confirm");
      localizeStub.withArgs("teamstoolkit.mcpUtils.setupMcpServer.skip").returns("Skip");
      localizeStub
        .withArgs("teamstoolkit.mcpUtils.setupMcpServer.successMessage")
        .returns("MCP Server setup completed successfully!");
      localizeStub
        .withArgs("teamstoolkit.mcpUtils.setupMcpServer.errorMessage")
        .returns("Error setting up MCP Server: %s");
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("should return early if no workspace folders", async () => {
      sandbox.stub(globalVariables, "workspaceUri").value(undefined);

      await setupMCPServer();

      chai.expect(showInformationMessageStub.called).to.be.false;
      chai.expect(sendTelemetryEventStub.called).to.be.false;
    });

    it("should return early if both files exist and MCP config is up to date", async () => {
      existsSyncStub.returns(true);

      const mockMcpConfig = {
        get: sandbox
          .stub()
          .withArgs("servers")
          .returns({
            "test-server": {
              command: "npx",
              args: ["@microsoft/m365agentstoolkit-mcp@latest", "server", "start"],
            },
          }),
      };
      getConfigurationStub.withArgs("mcp").returns(mockMcpConfig);

      await setupMCPServer();

      chai.expect(showInformationMessageStub.called).to.be.false;
      chai.expect(sendTelemetryEventStub.called).to.be.false;
    });

    it("should prompt user when copilot instructions file is missing", async () => {
      console.log(globalVariables.workspaceUri?.fsPath);
      existsSyncStub.callsFake((filePath: string) => {
        return filePath.endsWith("copilot-instructions.md") ? false : true;
      });
      const mockMcpConfig = {
        get: sandbox
          .stub()
          .withArgs("servers")
          .returns({
            "test-server": {
              command: "npx",
              args: ["@microsoft/m365agentstoolkit-mcp@latest", "server", "start"],
            },
          }),
      };
      getConfigurationStub.withArgs("mcp").returns(mockMcpConfig);

      showInformationMessageStub.resolves("Skip");

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
      existsSyncStub.callsFake((filePath: string) => {
        return filePath.endsWith("copilot-instructions.md") ? true : false;
      });

      const mockMcpConfig = {
        get: sandbox
          .stub()
          .withArgs("servers")
          .returns({
            "other-server": {
              command: "other-command",
            },
          }),
      };
      getConfigurationStub.withArgs("mcp").returns(mockMcpConfig);

      showInformationMessageStub.resolves("Skip");

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
      existsSyncStub.returns(false);

      const mockMcpConfig = {
        get: sandbox
          .stub()
          .withArgs("servers")
          .returns({
            "other-server": {
              command: "other-command",
            },
          }),
      };
      getConfigurationStub.withArgs("mcp").returns(mockMcpConfig);

      showInformationMessageStub.resolves("Skip");

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
      existsSyncStub.returns(false);
      const mkdirSyncStub = sandbox.stub(mcpUtilsDeps, "mkdirSync");
      const readFileSyncStub = sandbox.stub(mcpUtilsDeps, "readFileSync");
      const writeFileSyncStub = sandbox.stub(mcpUtilsDeps, "writeFileSync");
      const openSyncStub = sandbox.stub(mcpUtilsDeps, "openSync");
      sandbox.stub(mcpUtilsDeps, "closeSync");
      const showErrorMessageStub = sandbox
        .stub(mcpUtilsDeps, "showErrorMessage")
        .resolves(undefined);

      readFileSyncStub.returns("Default copilot instructions content");
      openSyncStub.returns(3);

      const mockMcpConfig = {
        get: sandbox
          .stub()
          .withArgs("servers")
          .returns({
            "other-server": {
              command: "other-command",
            },
          }),
      };
      getConfigurationStub.withArgs("mcp").returns(mockMcpConfig);

      showInformationMessageStub.resolves("Confirm");

      await setupMCPServer();

      chai.expect(mkdirSyncStub.calledTwice).to.be.true;
      chai.expect(writeFileSyncStub.calledTwice).to.be.true;
      chai.expect(showErrorMessageStub.called).to.be.false;

      chai.expect(
        sendTelemetryEventStub.calledWith(TelemetryEvent.PromptMCPServer, {
          [TelemetryProperty.UserSelection]: "confirm",
        })
      ).to.be.true;
    });

    it("should create only copilot instructions file when MCP config is up to date", async () => {
      existsSyncStub.callsFake((filePath: string) => {
        if (filePath.endsWith("copilot-instructions.md")) return false;
        if (filePath.endsWith(".github")) return false;
        return true;
      });

      const mkdirSyncStub = sandbox.stub(mcpUtilsDeps, "mkdirSync");
      const readFileSyncStub = sandbox.stub(mcpUtilsDeps, "readFileSync");
      const writeFileSyncStub = sandbox.stub(mcpUtilsDeps, "writeFileSync");
      const openSyncStub = sandbox.stub(mcpUtilsDeps, "openSync");
      sandbox.stub(mcpUtilsDeps, "closeSync");

      readFileSyncStub.returns("Default copilot instructions content");
      openSyncStub.returns(3);

      const mockMcpConfig = {
        get: sandbox
          .stub()
          .withArgs("servers")
          .returns({
            "test-server": {
              command: "npx",
              args: ["@microsoft/m365agentstoolkit-mcp@latest", "server", "start"],
            },
          }),
      };
      getConfigurationStub.withArgs("mcp").returns(mockMcpConfig);

      showInformationMessageStub.resolves("Confirm");

      await setupMCPServer();

      chai.expect(mkdirSyncStub.calledOnce).to.be.true;
      chai.expect(writeFileSyncStub.calledOnce).to.be.true;
      chai.expect(readFileSyncStub.calledOnce).to.be.true;
    });

    it("should create only MCP config file when copilot instructions exist", async () => {
      existsSyncStub.callsFake((filePath: string) => {
        if (filePath.endsWith("copilot-instructions.md")) return true;
        if (filePath.endsWith("mcp.json")) return false;
        if (filePath.endsWith(".vscode")) return false;
        return true;
      });

      const mkdirSyncStub = sandbox.stub(mcpUtilsDeps, "mkdirSync");
      const writeFileSyncStub = sandbox.stub(mcpUtilsDeps, "writeFileSync");

      const mockMcpConfig = {
        get: sandbox
          .stub()
          .withArgs("servers")
          .returns({
            "other-server": {
              command: "other-command",
            },
          }),
      };
      getConfigurationStub.withArgs("mcp").returns(mockMcpConfig);

      showInformationMessageStub.resolves("Confirm");

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
      existsSyncStub.callsFake((filePath: string) => {
        if (filePath.endsWith("copilot-instructions.md")) return true;
        if (filePath.endsWith("mcp.json")) return true;
        return true;
      });

      const readFileSyncStub = sandbox.stub(mcpUtilsDeps, "readFileSync");
      const writeFileSyncStub = sandbox.stub(mcpUtilsDeps, "writeFileSync");
      const openSyncStub = sandbox.stub(mcpUtilsDeps, "openSync");
      sandbox.stub(mcpUtilsDeps, "closeSync");

      readFileSyncStub.returns(
        JSON.stringify({
          servers: {
            "existing-server": {
              command: "existing-command",
            },
          },
        })
      );
      openSyncStub.returns(3);

      const mockMcpConfig = {
        get: sandbox
          .stub()
          .withArgs("servers")
          .returns({
            "other-server": {
              command: "other-command",
            },
          }),
      };
      getConfigurationStub.withArgs("mcp").returns(mockMcpConfig);

      showInformationMessageStub.resolves("Confirm");

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
      existsSyncStub.callsFake((filePath: string) => {
        if (filePath.endsWith("copilot-instructions.md")) return true;
        if (filePath.endsWith("mcp.json")) return true;
        return true;
      });

      const readFileSyncStub = sandbox.stub(mcpUtilsDeps, "readFileSync");
      const writeFileSyncStub = sandbox.stub(mcpUtilsDeps, "writeFileSync");
      const openSyncStub = sandbox.stub(mcpUtilsDeps, "openSync");
      sandbox.stub(mcpUtilsDeps, "closeSync");

      readFileSyncStub.returns(
        JSON.stringify({
          someOtherProperty: "value",
        })
      );
      openSyncStub.returns(3);

      const mockMcpConfig = {
        get: sandbox
          .stub()
          .withArgs("servers")
          .returns({
            "other-server": {
              command: "other-command",
            },
          }),
      };
      getConfigurationStub.withArgs("mcp").returns(mockMcpConfig);

      showInformationMessageStub.resolves("Confirm");

      await setupMCPServer();

      const writtenConfig = JSON.parse(writeFileSyncStub.firstCall.args[1] as string);
      chai.expect(writtenConfig.servers).to.exist;
      chai.expect(writtenConfig.servers["M365AgentsToolkit MCP Server"]).to.deep.equal({
        command: "npx",
        args: ["@microsoft/m365agentstoolkit-mcp@latest", "server", "start"],
      });
    });

    it("should handle MCP config file with null servers object", async () => {
      existsSyncStub.callsFake((filePath: string) => {
        if (filePath.endsWith("copilot-instructions.md")) return true;
        if (filePath.endsWith("mcp.json")) return true;
        return true;
      });

      const readFileSyncStub = sandbox.stub(mcpUtilsDeps, "readFileSync");
      const writeFileSyncStub = sandbox.stub(mcpUtilsDeps, "writeFileSync");
      const openSyncStub = sandbox.stub(mcpUtilsDeps, "openSync");
      sandbox.stub(mcpUtilsDeps, "closeSync");

      readFileSyncStub.returns(
        JSON.stringify({
          servers: null,
        })
      );
      openSyncStub.returns(3);

      const mockMcpConfig = {
        get: sandbox
          .stub()
          .withArgs("servers")
          .returns({
            "other-server": {
              command: "other-command",
            },
          }),
      };
      getConfigurationStub.withArgs("mcp").returns(mockMcpConfig);

      showInformationMessageStub.resolves("Confirm");

      await setupMCPServer();

      const writtenConfig = JSON.parse(writeFileSyncStub.firstCall.args[1] as string);
      chai.expect(writtenConfig.servers).to.exist;
      chai.expect(writtenConfig.servers["M365AgentsToolkit MCP Server"]).to.deep.equal({
        command: "npx",
        args: ["@microsoft/m365agentstoolkit-mcp@latest", "server", "start"],
      });
    });

    it("should handle error during file creation", async () => {
      existsSyncStub.returns(false);
      sandbox.stub(mcpUtilsDeps, "mkdirSync");
      const readFileSyncStub = sandbox.stub(mcpUtilsDeps, "readFileSync");
      const showErrorMessageStub = sandbox
        .stub(mcpUtilsDeps, "showErrorMessage")
        .resolves(undefined);
      const sendTelemetryErrorEventStub = sandbox.stub(mcpUtilsDeps, "sendTelemetryErrorEvent");

      const testError = new Error("Test error");
      readFileSyncStub.throws(testError);

      const mockMcpConfig = {
        get: sandbox
          .stub()
          .withArgs("servers")
          .returns({
            "other-server": {
              command: "other-command",
            },
          }),
      };
      getConfigurationStub.withArgs("mcp").returns(mockMcpConfig);

      showInformationMessageStub.resolves("Confirm");

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
      existsSyncStub.callsFake((filePath: string) => {
        return filePath.endsWith("copilot-instructions.md") ? true : false;
      });

      getConfigurationStub.withArgs("mcp").throws(new Error("Config error"));

      showInformationMessageStub.resolves("Skip");

      await setupMCPServer();

      chai.expect(
        sendTelemetryEventStub.calledWith(TelemetryEvent.PromptMCPServer, {
          [TelemetryProperty.MissingCopilotInstructions]: "false",
          [TelemetryProperty.MissingMCPConfig]: "true",
        })
      ).to.be.true;
    });

    it("should handle null servers configuration", async () => {
      existsSyncStub.callsFake((filePath: string) => {
        return filePath.endsWith("copilot-instructions.md") ? true : false;
      });

      const mockMcpConfig = {
        get: sandbox.stub().withArgs("servers").returns(null),
      };
      getConfigurationStub.withArgs("mcp").returns(mockMcpConfig);

      showInformationMessageStub.resolves("Skip");

      await setupMCPServer();

      chai.expect(
        sendTelemetryEventStub.calledWith(TelemetryEvent.PromptMCPServer, {
          [TelemetryProperty.MissingCopilotInstructions]: "false",
          [TelemetryProperty.MissingMCPConfig]: "true",
        })
      ).to.be.true;
    });

    it("should handle undefined servers configuration", async () => {
      existsSyncStub.callsFake((filePath: string) => {
        return filePath.endsWith("copilot-instructions.md") ? true : false;
      });

      const mockMcpConfig = {
        get: sandbox.stub().withArgs("servers").returns(undefined),
      };
      getConfigurationStub.withArgs("mcp").returns(mockMcpConfig);

      showInformationMessageStub.resolves("Skip");

      await setupMCPServer();

      chai.expect(
        sendTelemetryEventStub.calledWith(TelemetryEvent.PromptMCPServer, {
          [TelemetryProperty.MissingCopilotInstructions]: "false",
          [TelemetryProperty.MissingMCPConfig]: "true",
        })
      ).to.be.true;
    });

    it("should create github directory when it doesn't exist", async () => {
      existsSyncStub.callsFake((filePath: string) => {
        if (filePath.endsWith("copilot-instructions.md")) return false;
        if (filePath.endsWith(".github")) return false;
        return true;
      });

      const mkdirSyncStub = sandbox.stub(mcpUtilsDeps, "mkdirSync");
      const readFileSyncStub = sandbox.stub(mcpUtilsDeps, "readFileSync");
      const writeFileSyncStub = sandbox.stub(mcpUtilsDeps, "writeFileSync");
      const openSyncStub = sandbox.stub(mcpUtilsDeps, "openSync");
      sandbox.stub(mcpUtilsDeps, "closeSync");

      readFileSyncStub.returns("Default copilot instructions content");
      openSyncStub.returns(3);

      const mockMcpConfig = {
        get: sandbox
          .stub()
          .withArgs("servers")
          .returns({
            "test-server": {
              command: "npx",
              args: ["@microsoft/m365agentstoolkit-mcp@latest", "server", "start"],
            },
          }),
      };
      getConfigurationStub.withArgs("mcp").returns(mockMcpConfig);

      showInformationMessageStub.resolves("Confirm");

      await setupMCPServer();

      chai.expect(mkdirSyncStub.calledWith(path.join("/test", ".github"), { recursive: true })).to
        .be.true;
      chai.expect(writeFileSyncStub.calledOnce).to.be.true;
      chai.expect(readFileSyncStub.calledOnce).to.be.true;
      chai.expect(openSyncStub.calledOnce).to.be.true;
    });

    it("should create vscode directory when it doesn't exist", async () => {
      existsSyncStub.callsFake((filePath: string) => {
        if (filePath.endsWith("copilot-instructions.md")) return true;
        if (filePath.endsWith("mcp.json")) return false;
        if (filePath.endsWith(".vscode")) return false;
        return true;
      });

      const mkdirSyncStub = sandbox.stub(mcpUtilsDeps, "mkdirSync");
      const writeFileSyncStub = sandbox.stub(mcpUtilsDeps, "writeFileSync");

      const mockMcpConfig = {
        get: sandbox
          .stub()
          .withArgs("servers")
          .returns({
            "other-server": {
              command: "other-command",
            },
          }),
      };
      getConfigurationStub.withArgs("mcp").returns(mockMcpConfig);

      showInformationMessageStub.resolves("Confirm");

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
      existsSyncStub.returns(false);

      const mockMcpConfig = {
        get: sandbox
          .stub()
          .withArgs("servers")
          .returns({
            "other-server": {
              command: "other-command",
            },
          }),
      };
      getConfigurationStub.withArgs("mcp").returns(mockMcpConfig);

      showInformationMessageStub.resolves("Skip");

      await setupMCPServer();

      chai.expect(showInformationMessageStub.called).to.be.true;
      const messageCall = showInformationMessageStub.firstCall.args[0];
      chai.expect(messageCall).to.include("and");
    });
  });
});

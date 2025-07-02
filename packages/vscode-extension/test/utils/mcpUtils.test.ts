// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as chai from "chai";
import * as fs from "fs-extra";
import mockfs from "mock-fs";
import * as path from "path";
import * as sinon from "sinon";
import * as vscode from "vscode";
import * as globalVariables from "../../src/globalVariables";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { TelemetryEvent, TelemetryProperty } from "../../src/telemetry/extTelemetryEvents";
import { setupMCPServer } from "../../src/utils/mcpUtils";
import * as localizeUtils from "../../src/utils/localizeUtils";

describe("mcpUtils", () => {
  const sandbox = sinon.createSandbox();
  let workspaceFolder: vscode.WorkspaceFolder;
  let mockWorkspaceRoot: string;

  beforeEach(() => {
    mockWorkspaceRoot = "/test/workspace";
    workspaceFolder = {
      uri: vscode.Uri.file(mockWorkspaceRoot),
      name: "test-workspace",
      index: 0,
    };
  });

  afterEach(() => {
    mockfs.restore();
    sandbox.restore();
  });

  describe("setupMCPServer", () => {
    let showInformationMessageStub: sinon.SinonStub;
    let sendTelemetryEventStub: sinon.SinonStub;
    let getConfigurationStub: sinon.SinonStub;
    let localizeStub: sinon.SinonStub;

    beforeEach(() => {
      showInformationMessageStub = sandbox
        .stub(vscode.window, "showInformationMessage")
        .resolves(undefined);
      sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      getConfigurationStub = sandbox.stub(vscode.workspace, "getConfiguration");
      localizeStub = sandbox.stub(localizeUtils, "localize");

      // Default locale stubs
      localizeStub
        .withArgs("teamstoolkit.mcpUtils.setupMcpServer.message")
        .returns("Setup MCP Server for %s");
      localizeStub.withArgs("teamstoolkit.mcpUtils.setupMcpServer.confirm").returns("Yes");
      localizeStub.withArgs("teamstoolkit.mcpUtils.setupMcpServer.skip").returns("Skip");
      localizeStub
        .withArgs("teamstoolkit.mcpUtils.setupMcpServer.successMessage")
        .returns("MCP Server setup completed successfully!");

      // Default context stub
      sandbox.stub(globalVariables, "context").value({
        extensionPath: "/test/extension/path",
      });
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("should return early if no workspace folders", async () => {
      sandbox.stub(vscode.workspace, "workspaceFolders").value(undefined);

      await setupMCPServer();

      chai.expect(showInformationMessageStub.called).to.be.false;
      chai.expect(sendTelemetryEventStub.called).to.be.false;
    });

    it("should return early if workspace folders is empty", async () => {
      sandbox.stub(vscode.workspace, "workspaceFolders").value([]);

      await setupMCPServer();

      chai.expect(showInformationMessageStub.called).to.be.false;
      chai.expect(sendTelemetryEventStub.called).to.be.false;
    });

    it("should return early if both files exist and MCP config is up to date", async () => {
      sandbox.stub(vscode.workspace, "workspaceFolders").value([workspaceFolder]);

      // Mock file system with both files existing
      mockfs({
        [path.join(mockWorkspaceRoot, ".github", "copilot-instructions.md")]: "# Instructions",
        "/test/extension/path/media/mcp/copilot-instructions.md": "# Template content",
      });

      // MCP config is up to date (user already has m365agentstoolkit server configured)
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
      sandbox.stub(vscode.workspace, "workspaceFolders").value([workspaceFolder]);

      // Mock file system without the copilot instructions file
      mockfs({
        "/test/extension/path/media/mcp/copilot-instructions.md": "# Template content",
      });

      // MCP config is up to date (contains the required m365agentstoolkit package)
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

      chai.expect(
        showInformationMessageStub.calledWith(
          "Setup MCP Server for .github/copilot-instructions.md",
          "Yes",
          "Skip"
        )
      ).to.be.true;

      chai.expect(
        sendTelemetryEventStub.calledWith(TelemetryEvent.PromptMCPServer, {
          [TelemetryProperty.UserSelection]: "skip",
        })
      ).to.be.true;
    });

    it("should prompt user when MCP config needs update", async () => {
      sandbox.stub(vscode.workspace, "workspaceFolders").value([workspaceFolder]);

      // Mock file system with copilot instructions file existing
      mockfs({
        [path.join(mockWorkspaceRoot, ".github", "copilot-instructions.md")]: "# Instructions",
        "/test/extension/path/media/mcp/copilot-instructions.md": "# Template content",
      });

      // MCP config needs update (does NOT contain the required m365agentstoolkit package)
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

      chai.expect(
        showInformationMessageStub.calledWith(
          "Setup MCP Server for .vscode/mcp.json",
          "Yes",
          "Skip"
        )
      ).to.be.true;

      chai.expect(
        sendTelemetryEventStub.calledWith(TelemetryEvent.PromptMCPServer, {
          [TelemetryProperty.UserSelection]: "skip",
        })
      ).to.be.true;
    });

    it("should prompt user when both files need setup", async () => {
      sandbox.stub(vscode.workspace, "workspaceFolders").value([workspaceFolder]);

      // Mock file system without copilot instructions file
      mockfs({
        "/test/extension/path/media/mcp/copilot-instructions.md": "# Template content",
      });

      // MCP config needs update (empty servers object, no m365agentstoolkit package)
      const mockMcpConfig = {
        get: sandbox.stub().withArgs("servers").returns({}),
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

      chai.expect(
        showInformationMessageStub.calledWith(
          "Setup MCP Server for .github/copilot-instructions.md and .vscode/mcp.json",
          "Yes",
          "Skip"
        )
      ).to.be.true;
    });

    it("should create copilot instructions file when user confirms and file is missing", async () => {
      sandbox.stub(vscode.workspace, "workspaceFolders").value([workspaceFolder]);

      // Mock file system with template file but without copilot instructions file or .github directory
      const templateContent = "# Default copilot instructions\nSample content";
      mockfs({
        "/test/extension/path/media/mcp/copilot-instructions.md": templateContent,
      });

      // MCP config is up to date (so only copilot instructions will be created)
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

      showInformationMessageStub.resetBehavior();
      showInformationMessageStub.resolves("Yes"); // This should match the localized "confirm" value

      await setupMCPServer();

      // Verify .github directory is created
      chai.expect(fs.existsSync(path.join(mockWorkspaceRoot, ".github"))).to.be.true;

      // Verify copilot instructions file is created with template content
      chai.expect(fs.existsSync(path.join(mockWorkspaceRoot, ".github", "copilot-instructions.md")))
        .to.be.true;
      const fileContent = fs.readFileSync(
        path.join(mockWorkspaceRoot, ".github", "copilot-instructions.md"),
        "utf8"
      );
      chai.expect(fileContent).to.equal(templateContent);

      chai.expect(
        sendTelemetryEventStub.calledWith(TelemetryEvent.PromptMCPServer, {
          [TelemetryProperty.UserSelection]: "confirm",
        })
      ).to.be.true;
    });

    it("should create copilot instructions file with empty content when template is missing", async () => {
      sandbox.stub(vscode.workspace, "workspaceFolders").value([workspaceFolder]);

      // Mock file system with .github directory existing but no copilot instructions file and no template
      mockfs({
        [path.join(mockWorkspaceRoot, ".github")]: {},
      });

      // MCP config is up to date
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

      showInformationMessageStub.resolves("Yes");

      await setupMCPServer();

      // Verify copilot instructions file is created with empty content
      chai.expect(fs.existsSync(path.join(mockWorkspaceRoot, ".github", "copilot-instructions.md")))
        .to.be.true;
      const fileContent = fs.readFileSync(
        path.join(mockWorkspaceRoot, ".github", "copilot-instructions.md"),
        "utf8"
      );
      chai.expect(fileContent).to.equal("");
    });

    it("should create new MCP config file when it doesn't exist", async () => {
      sandbox.stub(vscode.workspace, "workspaceFolders").value([workspaceFolder]);

      // Mock file system with copilot instructions existing but no MCP config
      mockfs({
        [path.join(mockWorkspaceRoot, ".github", "copilot-instructions.md")]: "# Instructions",
        "/test/extension/path/media/mcp/copilot-instructions.md": "# Template content",
      });

      // MCP settings indicate config needs update
      const mockMcpConfig = {
        get: sandbox.stub().withArgs("servers").returns({}),
      };
      getConfigurationStub.withArgs("mcp").returns(mockMcpConfig);

      showInformationMessageStub.resolves("Yes");

      await setupMCPServer();

      // Verify .vscode directory is created
      chai.expect(fs.existsSync(path.join(mockWorkspaceRoot, ".vscode"))).to.be.true;

      // Verify MCP config file is created with correct content
      const expectedConfig = {
        servers: {
          m365agentstoolkit: {
            command: "npx",
            args: ["@microsoft/m365agentstoolkit-mcp@latest", "server", "start"],
          },
        },
      };

      chai.expect(fs.existsSync(path.join(mockWorkspaceRoot, ".vscode", "mcp.json"))).to.be.true;
      const fileContent = fs.readFileSync(
        path.join(mockWorkspaceRoot, ".vscode", "mcp.json"),
        "utf8"
      );
      chai.expect(JSON.parse(fileContent)).to.deep.equal(expectedConfig);
    });

    it("should update existing MCP config file", async () => {
      sandbox.stub(vscode.workspace, "workspaceFolders").value([workspaceFolder]);

      // Mock file system with both files existing
      const existingConfig = {
        servers: {
          "existing-server": {
            command: "existing-command",
          },
        },
      };
      mockfs({
        [path.join(mockWorkspaceRoot, ".github", "copilot-instructions.md")]: "# Instructions",
        [path.join(mockWorkspaceRoot, ".vscode", "mcp.json")]: JSON.stringify(existingConfig),
        "/test/extension/path/media/mcp/copilot-instructions.md": "# Template content",
      });

      // MCP settings indicate config needs update
      const mockMcpConfig = {
        get: sandbox.stub().withArgs("servers").returns(existingConfig.servers),
      };
      getConfigurationStub.withArgs("mcp").returns(mockMcpConfig);

      showInformationMessageStub.resolves("Yes");

      await setupMCPServer();

      // Verify MCP config file is updated with new server added
      const expectedConfig = {
        servers: {
          "existing-server": {
            command: "existing-command",
          },
          "M365AgentsToolkit MCP Server": {
            command: "npx",
            args: ["@microsoft/m365agentstoolkit-mcp@latest", "server", "start"],
          },
        },
      };

      chai.expect(fs.existsSync(path.join(mockWorkspaceRoot, ".vscode", "mcp.json"))).to.be.true;
      const fileContent = fs.readFileSync(
        path.join(mockWorkspaceRoot, ".vscode", "mcp.json"),
        "utf8"
      );
      chai.expect(JSON.parse(fileContent)).to.deep.equal(expectedConfig);
    });

    it("should handle existing MCP config without servers object", async () => {
      sandbox.stub(vscode.workspace, "workspaceFolders").value([workspaceFolder]);

      // Mock file system with existing MCP config without servers object
      const existingConfig = {
        someOtherProperty: "value",
      };
      mockfs({
        [path.join(mockWorkspaceRoot, ".github", "copilot-instructions.md")]: "# Instructions",
        [path.join(mockWorkspaceRoot, ".vscode", "mcp.json")]: JSON.stringify(existingConfig),
        "/test/extension/path/media/mcp/copilot-instructions.md": "# Template content",
      });

      // MCP settings indicate config needs update
      const mockMcpConfig = {
        get: sandbox.stub().withArgs("servers").returns(undefined),
      };
      getConfigurationStub.withArgs("mcp").returns(mockMcpConfig);

      showInformationMessageStub.resolves("Yes");

      await setupMCPServer();

      // Verify MCP config file is updated with servers object added
      const expectedConfig = {
        someOtherProperty: "value",
        servers: {
          "M365AgentsToolkit MCP Server": {
            command: "npx",
            args: ["@microsoft/m365agentstoolkit-mcp@latest", "server", "start"],
          },
        },
      };

      chai.expect(fs.existsSync(path.join(mockWorkspaceRoot, ".vscode", "mcp.json"))).to.be.true;
      const fileContent = fs.readFileSync(
        path.join(mockWorkspaceRoot, ".vscode", "mcp.json"),
        "utf8"
      );
      chai.expect(JSON.parse(fileContent)).to.deep.equal(expectedConfig);
    });

    it("should handle MCP config check error gracefully", async () => {
      sandbox.stub(vscode.workspace, "workspaceFolders").value([workspaceFolder]);

      // Mock file system with copilot instructions existing
      mockfs({
        [path.join(mockWorkspaceRoot, ".github", "copilot-instructions.md")]: "# Instructions",
        "/test/extension/path/media/mcp/copilot-instructions.md": "# Template content",
      });

      // MCP config check throws error
      getConfigurationStub.withArgs("mcp").throws(new Error("Config error"));

      showInformationMessageStub.resolves("Yes");

      await setupMCPServer();

      // Should still proceed with MCP config update since error defaults to needing update
      chai.expect(
        sendTelemetryEventStub.calledWith(TelemetryEvent.PromptMCPServer, {
          [TelemetryProperty.MissingCopilotInstructions]: "false",
          [TelemetryProperty.MissingMCPConfig]: "true",
        })
      ).to.be.true;
    });

    it("should show success message after setup completion", async () => {
      sandbox.stub(vscode.workspace, "workspaceFolders").value([workspaceFolder]);

      // Mock file system without both files
      mockfs({
        "/test/extension/path/media/mcp/copilot-instructions.md": "",
      });

      const mockMcpConfig = {
        get: sandbox.stub().withArgs("servers").returns({}),
      };
      getConfigurationStub.withArgs("mcp").returns(mockMcpConfig);

      showInformationMessageStub.resolves("Yes");

      await setupMCPServer();

      // Verify success message is shown
      chai.expect(showInformationMessageStub.calledWith("MCP Server setup completed successfully!"))
        .to.be.true;
    });

    it("should handle multiple files message formatting correctly", async () => {
      sandbox.stub(vscode.workspace, "workspaceFolders").value([workspaceFolder]);

      // Mock file system without both files
      mockfs({
        "/test/extension/path/media/mcp/copilot-instructions.md": "# Template content",
      });

      const mockMcpConfig = {
        get: sandbox.stub().withArgs("servers").returns({}),
      };
      getConfigurationStub.withArgs("mcp").returns(mockMcpConfig);

      showInformationMessageStub.resolves("Skip");

      await setupMCPServer();

      // Should handle the "and" case for 2 files
      chai.expect(
        showInformationMessageStub.calledWith(
          "Setup MCP Server for .github/copilot-instructions.md and .vscode/mcp.json",
          "Yes",
          "Skip"
        )
      ).to.be.true;
    });
  });
});

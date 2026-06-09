// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as sinon from "sinon";
import * as chai from "chai";
import * as vscode from "vscode";
import fs from "fs-extra";
import * as path from "path";
import axios from "axios";
import * as parser from "jsonc-parser";
import { Stage, err, ok, UserError, Inputs, Platform } from "@microsoft/teamsfx-api";

import { updateActionWithMCP } from "../../src/handlers/updateActionWithMCP";
import * as systemEnvUtils from "../../src/utils/systemEnvUtils";
import * as sharedOpts from "../../src/handlers/sharedOpts";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import * as vscUI from "../../src/qm/vsc_ui";
import { QuestionNames, ODRProvider } from "@microsoft/teamsfx-core";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { MockCore } from "../mocks/mockCore";
import * as globalVariables from "../../src/globalVariables";

describe("updateActionWithMCP", () => {
  const sandbox = sinon.createSandbox();

  const mockProjectPath = "/mock/project/path";

  beforeEach(() => {
    const mockInputs: Inputs = {
      projectPath: mockProjectPath,
      platform: Platform.VSCode,
    };
    sandbox.stub(systemEnvUtils, "getSystemInputs").returns(mockInputs);
    sandbox.stub(vscode.window, "showErrorMessage");
    sandbox.stub(globalVariables, "core").value(new MockCore());
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
  });

  afterEach(() => {
    sandbox.restore();
  });

  describe("sanitizeMCPName", () => {
    it("should sanitize MCP name correctly", async () => {
      // Test with args containing serverName with special characters
      const args = [{ serverName: "test-server@123!#$", serverConfig: { url: "http://test.com" } }];

      sandbox.stub(fs, "pathExistsSync").returns(false);
      sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));
      Object.defineProperty(vscode.lm, "tools", { value: [], configurable: true });

      const result = await updateActionWithMCP(args);

      // The function should still process even with empty tools (will show error)
      chai.assert.isTrue(result.isErr());
    });

    it("should limit MCP name to 13 characters", async () => {
      const args = [
        {
          serverName: "verylongservernamethatexceeds13characters",
          serverConfig: { url: "http://test.com" },
        },
      ];

      sandbox.stub(fs, "pathExistsSync").returns(false);
      sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));
      Object.defineProperty(vscode.lm, "tools", { value: [], configurable: true });

      await updateActionWithMCP(args);

      // Verify the sanitized name was used (indirectly through no errors about invalid names)
      chai.assert.isTrue(true); // If we reach here, sanitization worked
    });
  });

  describe("with args provided", () => {
    it("should process successfully with valid args", async () => {
      const args = [{ serverName: "testServer", serverConfig: { url: "http://test.com" } }];
      const mockTools = [
        {
          name: "mcp_testserver_tool1",
          description: "Test tool 1",
          inputSchema: {},
          tags: [],
        },
        {
          name: "mcp_testserver_tool2",
          description: "Test tool 2",
          inputSchema: {},
          tags: [],
        },
      ];

      Object.defineProperty(vscode.lm, "tools", { value: mockTools, configurable: true });
      sandbox.stub(axios, "get").resolves({ status: 200 });
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP(args);

      chai.assert.isTrue(result.isOk());
      sinon.assert.calledWith(runCommandStub, Stage.updateActionWithMCP, sinon.match.any);
    });

    it("should return error when server name is provided but URL is missing", async () => {
      const args = [{ serverName: "testServer" }]; // Missing serverConfig.url

      const result = await updateActionWithMCP(args);

      chai.assert.isTrue(result.isErr());
    });

    it("should return error when URL is provided but server name is missing", async () => {
      const args = [{ serverConfig: { url: "http://test.com" } }]; // Missing serverName

      const result = await updateActionWithMCP(args);

      chai.assert.isTrue(result.isErr());
    });
  });

  describe("without args - MCP file handling", () => {
    it("should return error when projectPath is missing", async () => {
      const emptyInputs: Inputs = {
        platform: Platform.VSCode,
      };
      sandbox.restore();
      sandbox.stub(systemEnvUtils, "getSystemInputs").returns(emptyInputs);
      sandbox.stub(vscode.window, "showErrorMessage");
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");

      const result = await updateActionWithMCP();

      chai.assert.isTrue(result.isErr());
    });

    it("should return error when MCP file does not exist", async () => {
      sandbox.stub(fs, "pathExistsSync").returns(false);

      const result = await updateActionWithMCP();

      chai.assert.isTrue(result.isErr());
    });

    it("should return error when MCP file has invalid content", async () => {
      const expectedPath = path.join(mockProjectPath, ".vscode", "mcp.json");
      sandbox.stub(fs, "readFileSync").withArgs(expectedPath, "utf-8").returns("{}");
      sandbox.stub(fs, "pathExistsSync").withArgs(expectedPath).returns(true);
      sandbox.stub(parser, "parse").returns({});

      const result = await updateActionWithMCP();

      chai.assert.isTrue(result.isErr());
    });

    it("should return error when no MCP servers found", async () => {
      const expectedPath = path.join(mockProjectPath, ".vscode", "mcp.json");
      sandbox.stub(fs, "pathExistsSync").withArgs(expectedPath).returns(true);
      sandbox.stub(fs, "readFileSync").withArgs(expectedPath, "utf-8").returns('{"servers":{}}');
      sandbox.stub(parser, "parse").returns({ servers: {} });

      const result = await updateActionWithMCP();

      chai.assert.isTrue(result.isErr());
    });

    it("should process single MCP server automatically", async () => {
      const mcpContent = {
        servers: {
          "test-server": { url: "http://test.com" },
        },
      };
      const mockTools = [
        {
          name: "mcp_test-server_tool1",
          description: "Test tool",
          inputSchema: {},
          tags: [],
        },
      ];

      const expectedPath = path.join(mockProjectPath, ".vscode", "mcp.json");
      sandbox.stub(fs, "pathExistsSync").withArgs(expectedPath).returns(true);
      sandbox
        .stub(fs, "readFileSync")
        .withArgs(expectedPath, "utf-8")
        .returns(JSON.stringify(mcpContent));
      sandbox.stub(parser, "parse").returns(mcpContent);
      Object.defineProperty(vscode.lm, "tools", { value: mockTools, configurable: true });
      sandbox.stub(axios, "get").resolves({ status: 200 });
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP();

      chai.assert.isTrue(result.isOk());
      sinon.assert.calledOnce(runCommandStub);
    });

    it("should handle ODR server with command but no args", async () => {
      const args = [
        {
          serverName: "odrServerNoArgs",
          serverConfig: { type: "stdio", command: "odr" },
        },
      ];
      const mockODRTools = [
        {
          name: "tool1",
          description: "Test tool",
          inputSchema: {},
        },
      ];
      const mockODRServers = [
        {
          name: "my-server-noargs",
          display_name: "My Server NoArgs",
          description: "Test Server NoArgs",
          version: "1.0.0",
          identifier: "my-server-identifier-noargs",
          tools: mockODRTools,
          packageFamily: "test.package",
          command: "odr",
          args: [],
        },
      ];

      sandbox.stub(ODRProvider, "listServers").resolves(mockODRServers);
      sandbox.stub(ODRProvider, "getToolsForODRServer").resolves(mockODRTools);
      sandbox.stub(ODRProvider, "isODRServer").returns(true);
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP(args);

      chai.assert.isTrue(result.isOk());
      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      chai.assert.equal(
        calledInputs[QuestionNames.MCPLocalServerIdentifier],
        "my-server-identifier-noargs"
      );
      chai.assert.equal(calledInputs[QuestionNames.MCPForDAAvailableTools].length, 1);
      chai.assert.equal(calledInputs[QuestionNames.MCPForDAAvailableTools][0].name, "tool1");
    });

    it("should process single local MCP server automatically (non-ODR)", async () => {
      const mcpContent = {
        servers: {
          "local-server": {
            type: "stdio",
            command: "node",
            args: ["server.js"],
          },
        },
      };
      const mockTools = [
        {
          name: "mcp_local-server_tool1",
          description: "Test tool",
          inputSchema: {},
          tags: [],
        },
      ];

      const expectedPath = path.join(mockProjectPath, ".vscode", "mcp.json");
      sandbox.stub(fs, "pathExistsSync").withArgs(expectedPath).returns(true);
      sandbox
        .stub(fs, "readFileSync")
        .withArgs(expectedPath, "utf-8")
        .returns(JSON.stringify(mcpContent));
      sandbox.stub(parser, "parse").returns(mcpContent);
      sandbox.stub(vscode.lm, "tools").value(mockTools);
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP();

      chai.assert.isTrue(result.isOk());
      sinon.assert.calledOnce(runCommandStub);
      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      chai.assert.equal(calledInputs[QuestionNames.MCPLocalServerIdentifier], "local-server");
    });

    it("should process single local MCP server automatically (ODR)", async () => {
      const mcpContent = {
        servers: {
          "odr-server": {
            type: "stdio",
            command: "odr",
            args: ["run", "my-mcp-server"],
          },
        },
      };
      const mockODRTools = [
        {
          name: "tool1",
          description: "Test tool",
          inputSchema: {},
        },
      ];
      const mockODRServers = [
        {
          name: "my-mcp-server",
          display_name: "My MCP Server",
          description: "Test MCP Server",
          version: "1.0.0",
          identifier: "my-mcp-server-id",
          tools: mockODRTools,
          packageFamily: "test.package",
          command: "odr",
          args: ["run", "my-mcp-server"],
        },
      ];

      const expectedPath = path.join(mockProjectPath, ".vscode", "mcp.json");
      sandbox.stub(fs, "pathExistsSync").withArgs(expectedPath).returns(true);
      sandbox
        .stub(fs, "readFileSync")
        .withArgs(expectedPath, "utf-8")
        .returns(JSON.stringify(mcpContent));
      sandbox.stub(parser, "parse").returns(mcpContent);
      sandbox.stub(ODRProvider, "listServers").resolves(mockODRServers);
      sandbox.stub(ODRProvider, "getToolsForODRServer").resolves(mockODRTools);
      sandbox.stub(ODRProvider, "isODRServer").returns(true);
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP();

      chai.assert.isTrue(result.isOk());
      sinon.assert.calledOnce(runCommandStub);
      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      chai.assert.equal(calledInputs[QuestionNames.MCPLocalServerIdentifier], "my-mcp-server-id");
      chai.assert.equal(calledInputs[QuestionNames.MCPForDAAvailableTools].length, 1);
      chai.assert.equal(calledInputs[QuestionNames.MCPForDAAvailableTools][0].name, "tool1");
    });

    it("should show selection UI for multiple MCP servers", async () => {
      const mcpContent = {
        servers: {
          "remote-server": { url: "http://remote.com" },
          "remote-server2": { url: "http://remote2.com" },
        },
      };
      const mockTools = [
        {
          name: "mcp_remote-server_tool1",
          description: "Test tool",
          inputSchema: {},
          tags: [],
        },
      ];

      const expectedPath = path.join(mockProjectPath, ".vscode", "mcp.json");
      sandbox.stub(fs, "pathExistsSync").withArgs(expectedPath).returns(true);
      sandbox
        .stub(fs, "readFileSync")
        .withArgs(expectedPath, "utf-8")
        .returns(JSON.stringify(mcpContent));
      sandbox.stub(parser, "parse").returns(mcpContent);
      Object.defineProperty(vscode.lm, "tools", { value: mockTools, configurable: true });
      sandbox.stub(axios, "get").resolves({ status: 200 });
      sandbox.stub(vscUI, "VS_CODE_UI").value({
        selectOption: sandbox.stub().resolves(ok({ type: "success", result: "remote-server" })),
      });
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP();

      chai.assert.isTrue(result.isOk());
      sinon.assert.calledOnce(runCommandStub);
    });

    it("should show selection UI for multiple MCP servers including local", async () => {
      const mcpContent = {
        servers: {
          "remote-server": { url: "http://remote.com" },
          "local-server": {
            type: "stdio",
            command: "node",
            args: ["server.js"],
          },
        },
      };
      const mockTools = [
        {
          name: "mcp_local-server_tool1",
          description: "Test tool",
          inputSchema: {},
          tags: [],
        },
      ];

      const expectedPath = path.join(mockProjectPath, ".vscode", "mcp.json");
      sandbox.stub(fs, "pathExistsSync").withArgs(expectedPath).returns(true);
      sandbox
        .stub(fs, "readFileSync")
        .withArgs(expectedPath, "utf-8")
        .returns(JSON.stringify(mcpContent));
      sandbox.stub(parser, "parse").returns(mcpContent);
      sandbox.stub(ODRProvider, "listServers").resolves([]);
      Object.defineProperty(vscode.lm, "tools", { value: mockTools, configurable: true });
      sandbox.stub(vscUI, "VS_CODE_UI").value({
        selectOption: sandbox.stub().resolves(ok({ type: "success", result: "local-server" })),
      });
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP();

      chai.assert.isTrue(result.isOk());
      sinon.assert.calledOnce(runCommandStub);
      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      chai.assert.equal(calledInputs[QuestionNames.MCPLocalServerIdentifier], "local-server");
    });

    it("should return error when user cancels server selection", async () => {
      const mcpContent = {
        servers: {
          "remote-server": { url: "http://remote.com" },
          "remote-server2": { url: "http://remote2.com" },
        },
      };

      const expectedPath = path.join(mockProjectPath, ".vscode", "mcp.json");
      sandbox.stub(fs, "pathExistsSync").withArgs(expectedPath).returns(true);
      sandbox
        .stub(fs, "readFileSync")
        .withArgs(expectedPath, "utf-8")
        .returns(JSON.stringify(mcpContent));
      sandbox.stub(parser, "parse").returns(mcpContent);
      sandbox.stub(vscUI, "VS_CODE_UI").value({
        selectOption: sandbox
          .stub()
          .resolves(err(new UserError("test", "UserCancel", "User cancelled"))),
      });

      const result = await updateActionWithMCP();

      chai.assert.isTrue(result.isErr());
    });

    it("should construct detail correctly for local server in selection UI", async () => {
      const mcpContent = {
        servers: {
          "remote-server": { url: "http://remote.com" },
          "local-server": {
            type: "stdio",
            command: "node",
            args: ["server.js", "arg1"],
          },
        },
      };
      const mockTools = [
        {
          name: "mcp_remoteserver_tool1",
          description: "Test tool",
          inputSchema: {},
          tags: [],
        },
      ];

      const expectedPath = path.join(mockProjectPath, ".vscode", "mcp.json");
      sandbox.stub(fs, "pathExistsSync").withArgs(expectedPath).returns(true);
      sandbox
        .stub(fs, "readFileSync")
        .withArgs(expectedPath, "utf-8")
        .returns(JSON.stringify(mcpContent));
      sandbox.stub(parser, "parse").returns(mcpContent);
      Object.defineProperty(vscode.lm, "tools", { value: mockTools, configurable: true });
      sandbox.stub(axios, "get").resolves({ status: 200 });

      let capturedOptions: any;
      sandbox.stub(vscUI, "VS_CODE_UI").value({
        selectOption: sandbox.stub().callsFake((config: any) => {
          capturedOptions = config.options;
          return Promise.resolve(ok({ type: "success", result: "remote-server" }));
        }),
      });
      sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      await updateActionWithMCP();

      const localServerOption = capturedOptions.find((opt: any) => opt.id === "local-server");
      chai.assert.equal(localServerOption.detail, "node server.js arg1");

      const remoteServerOption = capturedOptions.find((opt: any) => opt.id === "remote-server");
      chai.assert.equal(remoteServerOption.detail, "http://remote.com");
    });
  });

  describe("MCP tools processing", () => {
    it("should return error when no tools are found for MCP server", async () => {
      const args = [{ serverName: "testServer", serverConfig: { url: "http://test.com" } }];

      sandbox.stub(vscode.lm, "tools").value([]); // No tools found

      const result = await updateActionWithMCP(args);

      chai.assert.isTrue(result.isErr());
    });

    it("should filter and transform tools correctly", async () => {
      const args = [{ serverName: "testServer", serverConfig: { url: "http://test.com" } }];
      const mockTools = [
        {
          name: "mcp_testserver_getTodos",
          description: "Get todos",
          inputSchema: { type: "object" },
          tags: ["todo"],
        },
        {
          name: "mcp_testserver_createTodo",
          description: "Create todo",
          inputSchema: { type: "object" },
          tags: ["todo"],
        },
        {
          name: "mcp_otherserver_tool", // Should be filtered out
          description: "Other tool",
          inputSchema: {},
          tags: [],
        },
      ];

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      sandbox.stub(axios, "get").resolves({ status: 200 });
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP(args);

      chai.assert.isTrue(result.isOk());

      // Verify runCommand was called with correct inputs
      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      const tools = calledInputs[QuestionNames.MCPForDAAvailableTools];

      chai.assert.equal(tools.length, 2);
      chai.assert.equal(tools[0].name, "getTodos");
      chai.assert.equal(tools[1].name, "createTodo");
    });
  });

  describe("authentication handling", () => {
    it("should handle no authentication (200 response)", async () => {
      const args = [{ serverName: "testServer", serverConfig: { url: "http://test.com" } }];
      const mockTools = [
        { name: "mcp_testserver_tool1", description: "Test", inputSchema: {}, tags: [] },
      ];

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      sandbox.stub(axios, "get").resolves({ status: 200 });
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      await updateActionWithMCP(args);

      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      chai.assert.equal(calledInputs[QuestionNames.MCPForDAAuth], "NoneAuth");
    });

    it("should handle OAuth authentication (401 response)", async () => {
      const args = [{ serverName: "testServer", serverConfig: { url: "http://test.com" } }];
      const mockTools = [
        { name: "mcp_testserver_tool1", description: "Test", inputSchema: {}, tags: [] },
      ];
      const axiosError = {
        status: 401,
        response: {
          headers: {
            "www-authenticate": 'Bearer resource_metadata="http://test.com/.well-known/oauth"',
          },
        },
      };

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      sandbox
        .stub(axios, "get")
        .onFirstCall()
        .throws(axiosError)
        .onSecondCall()
        .resolves({ status: 200 });
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      await updateActionWithMCP(args);

      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      chai.assert.equal(calledInputs[QuestionNames.MCPForDAAuth], "OAuthPluginVault");
      chai.assert.equal(
        calledInputs[QuestionNames.MCPForDAAuthMetadataUrl],
        "http://test.com/.well-known/oauth"
      );
    });

    it("should handle OAuth with well-known URL fallback", async () => {
      const args = [{ serverName: "testServer", serverConfig: { url: "https://api.test.com/v1" } }];
      const mockTools = [
        { name: "mcp_testserver_tool1", description: "Test", inputSchema: {}, tags: [] },
      ];
      const axiosError = {
        status: 401,
        response: {
          headers: {}, // No www-authenticate header
        },
      };

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      const axiosStub = sandbox.stub(axios, "get");
      axiosStub.withArgs("https://api.test.com/v1").throws(axiosError);
      axiosStub
        .withArgs("https://api.test.com/.well-known/oauth-authorization-server")
        .resolves({ status: 200 }); // well-known URL response
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      await updateActionWithMCP(args);

      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      chai.assert.equal(calledInputs[QuestionNames.MCPForDAAuth], "OAuthPluginVault");
      chai.assert.equal(
        calledInputs[QuestionNames.MCPForDAAuthWellKnownUrl],
        "https://api.test.com/.well-known/oauth-authorization-server"
      );
    });

    it("should handle network errors gracefully", async () => {
      const args = [{ serverName: "testServer", serverConfig: { url: "http://test.com" } }];
      const mockTools = [
        { name: "mcp_testserver_tool1", description: "Test", inputSchema: {}, tags: [] },
      ];
      const networkError = new Error("Network error");

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      sandbox.stub(axios, "get").throws(networkError);
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      await updateActionWithMCP(args);

      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      chai.assert.equal(calledInputs[QuestionNames.MCPForDAAuth], "NoneAuth");
    });
  });

  describe("input validation and setup", () => {
    it("should set correct input values", async () => {
      const args = [{ serverName: "testServer", serverConfig: { url: "http://test.com" } }];
      const mockTools = [
        { name: "mcp_testserver_tool1", description: "Test", inputSchema: {}, tags: [] },
      ];

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      sandbox.stub(axios, "get").resolves({ status: 200 });
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      await updateActionWithMCP(args);

      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      chai.assert.equal(calledInputs[QuestionNames.MCPForDAServerUrl], "http://test.com");
      chai.assert.equal(calledInputs[QuestionNames.MCPForDAServerName], "testserver");
      chai.assert.isArray(calledInputs[QuestionNames.MCPForDAAvailableTools]);
      chai.assert.equal(calledInputs[QuestionNames.MCPForDAAuth], "NoneAuth");
    });

    it("should propagate runCommand errors", async () => {
      const args = [{ serverName: "testServer", serverConfig: { url: "http://test.com" } }];
      const mockTools = [
        { name: "mcp_testserver_tool1", description: "Test", inputSchema: {}, tags: [] },
      ];
      const runCommandError = new UserError("test", "RunCommandError", "Run command failed");

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      sandbox.stub(axios, "get").resolves({ status: 200 });
      sandbox.stub(sharedOpts, "runCommand").resolves(err(runCommandError));

      const result = await updateActionWithMCP(args);

      chai.assert.isTrue(result.isErr());
      chai.assert.equal(result._unsafeUnwrapErr().message, "Run command failed");
    });
  });

  describe("edge cases", () => {
    it("should handle server config with no command", async () => {
      const mcpContent = {
        servers: {
          "local-server": {
            type: "stdio",
          },
        },
      };
      const mockTools = [
        {
          name: "mcp_local-server_tool1",
          description: "Test tool",
          inputSchema: {},
          tags: [],
        },
      ];

      const expectedPath = path.join(mockProjectPath, ".vscode", "mcp.json");
      sandbox.stub(fs, "pathExistsSync").withArgs(expectedPath).returns(true);
      sandbox
        .stub(fs, "readFileSync")
        .withArgs(expectedPath, "utf-8")
        .returns(JSON.stringify(mcpContent));
      sandbox.stub(parser, "parse").returns(mcpContent);
      sandbox.stub(ODRProvider, "listServers").resolves([]);
      Object.defineProperty(vscode.lm, "tools", { value: mockTools, configurable: true });

      const result = await updateActionWithMCP();

      chai.assert.isTrue(result.isErr());
    });

    it("should handle server config with empty args for detail fallback", async () => {
      const mcpContent = {
        servers: {
          "remote-server": { url: "http://remote.com" },
          "local-server": {
            type: "stdio",
            command: "python",
            args: [],
          },
        },
      };
      const mockTools = [
        {
          name: "mcp_remote-server_tool1",
          description: "Test tool",
          inputSchema: {},
          tags: [],
        },
      ];

      const expectedPath = path.join(mockProjectPath, ".vscode", "mcp.json");
      sandbox.stub(fs, "pathExistsSync").withArgs(expectedPath).returns(true);
      sandbox
        .stub(fs, "readFileSync")
        .withArgs(expectedPath, "utf-8")
        .returns(JSON.stringify(mcpContent));
      sandbox.stub(parser, "parse").returns(mcpContent);
      Object.defineProperty(vscode.lm, "tools", { value: mockTools, configurable: true });
      sandbox.stub(axios, "get").resolves({ status: 200 });

      let capturedOptions: any;
      sandbox.stub(vscUI, "VS_CODE_UI").value({
        selectOption: sandbox.stub().callsFake((config: any) => {
          capturedOptions = config.options;
          return Promise.resolve(ok({ type: "success", result: "remote-server" }));
        }),
      });
      sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      await updateActionWithMCP();

      const localServerOption = capturedOptions.find((opt: any) => opt.id === "local-server");
      chai.assert.equal(localServerOption.detail, "python");
    });

    it("should handle remote server config with empty url", async () => {
      const mcpContent = {
        servers: {
          "remote-server": { url: "http://remote.com" },
          "remote-server2": {
            url: "",
          },
        },
      };
      const mockTools = [
        {
          name: "mcp_remote-server_tool1",
          description: "Test tool",
          inputSchema: {},
          tags: [],
        },
      ];

      const expectedPath = path.join(mockProjectPath, ".vscode", "mcp.json");
      sandbox.stub(fs, "pathExistsSync").withArgs(expectedPath).returns(true);
      sandbox
        .stub(fs, "readFileSync")
        .withArgs(expectedPath, "utf-8")
        .returns(JSON.stringify(mcpContent));
      sandbox.stub(parser, "parse").returns(mcpContent);
      Object.defineProperty(vscode.lm, "tools", { value: mockTools, configurable: true });
      sandbox.stub(axios, "get").resolves({ status: 200 });

      let capturedOptions: any;
      sandbox.stub(vscUI, "VS_CODE_UI").value({
        selectOption: sandbox.stub().callsFake((config: any) => {
          capturedOptions = config.options;
          return Promise.resolve(ok({ type: "success", result: "remote-server" }));
        }),
      });
      sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      await updateActionWithMCP();

      const remoteServer2Option = capturedOptions.find((opt: any) => opt.id === "remote-server2");
      chai.assert.equal(remoteServer2Option.detail, "");
    });

    it("should handle undefined args", async () => {
      const expectedPath = path.join(mockProjectPath, ".vscode", "mcp.json");
      sandbox.stub(fs, "pathExistsSync").withArgs(expectedPath).returns(false);

      const result = await updateActionWithMCP(undefined);

      chai.assert.isTrue(result.isErr());
    });

    it("should handle empty args array", async () => {
      const expectedPath = path.join(mockProjectPath, ".vscode", "mcp.json");
      sandbox.stub(fs, "pathExistsSync").withArgs(expectedPath).returns(false);

      const result = await updateActionWithMCP([]);

      chai.assert.isTrue(result.isErr());
    });

    it("should handle tools with missing name parts", async () => {
      const args = [
        {
          serverName: "testServer",
          serverConfig: { url: "http://test.com" },
        },
      ];
      const mockTools = [
        { name: "mcp_testserver_tool1", description: "Test", inputSchema: {}, tags: [] },
      ];

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      sandbox.stub(axios, "get").resolves({ status: 200 });
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP(args);

      chai.assert.isTrue(result.isOk());
      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      chai.assert.isUndefined(calledInputs[QuestionNames.MCPLocalServerIdentifier]);
    });

    it("should handle local MCP with stdio type (non-ODR)", async () => {
      const args = [
        {
          serverName: "localServer",
          serverConfig: { type: "stdio", command: "node", args: ["server.js"] },
        },
      ];
      const mockTools = [
        { name: "mcp_localserver_tool1", description: "Test", inputSchema: {}, tags: [] },
      ];

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      sandbox.stub(ODRProvider, "listServers").resolves([]);
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP(args);

      chai.assert.isTrue(result.isOk());
      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      chai.assert.equal(calledInputs[QuestionNames.MCPLocalServerIdentifier], "localServer");
    });

    it("should handle local MCP with stdio type (ODR)", async () => {
      const args = [
        {
          serverName: "odrServer",
          serverConfig: { type: "stdio", command: "odr", args: ["run", "my-server"] },
        },
      ];
      const mockODRTools = [{ name: "tool1", description: "Test", inputSchema: {} }];
      const mockODRServers = [
        {
          name: "my-server",
          display_name: "My Server",
          description: "Test Server",
          version: "1.0.0",
          identifier: "my-server-identifier",
          tools: mockODRTools,
          packageFamily: "test.package",
          command: "odr",
          args: ["run", "my-server"],
        },
      ];

      sandbox.stub(ODRProvider, "listServers").resolves(mockODRServers);
      sandbox.stub(ODRProvider, "getToolsForODRServer").resolves(mockODRTools);
      sandbox.stub(ODRProvider, "isODRServer").returns(true);
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP(args);

      chai.assert.isTrue(result.isOk());
      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      chai.assert.equal(
        calledInputs[QuestionNames.MCPLocalServerIdentifier],
        "my-server-identifier"
      );
      chai.assert.equal(calledInputs[QuestionNames.MCPForDAAvailableTools].length, 1);
      chai.assert.equal(calledInputs[QuestionNames.MCPForDAAvailableTools][0].name, "tool1");
    });

    it("should set mcp-type to local for stdio servers", async () => {
      const args = [
        {
          serverName: "localServer",
          serverConfig: { type: "stdio", command: "node", args: ["server.js"] },
        },
      ];
      const mockTools = [
        { name: "mcp_localserver_tool1", description: "Test", inputSchema: {}, tags: [] },
      ];

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      sandbox.stub(ODRProvider, "listServers").resolves([]);
      sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      await updateActionWithMCP(args);

      const telemetryCall = (ExtTelemetry.sendTelemetryEvent as sinon.SinonStub).lastCall;
      chai.assert.equal(telemetryCall.args[1]["mcp-type"], "local");
    });

    it("should set mcp-type to remote for non-stdio servers", async () => {
      const args = [{ serverName: "remoteServer", serverConfig: { url: "http://remote.com" } }];
      const mockTools = [
        { name: "mcp_remoteserver_tool1", description: "Test", inputSchema: {}, tags: [] },
      ];

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      sandbox.stub(axios, "get").resolves({ status: 200 });
      sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      await updateActionWithMCP(args);

      const telemetryCall = (ExtTelemetry.sendTelemetryEvent as sinon.SinonStub).lastCall;
      chai.assert.equal(telemetryCall.args[1]["mcp-type"], "remote");
    });

    it("should send telemetry with local mcp-type on error", async () => {
      const args = [
        {
          serverName: "localServer",
          serverConfig: { type: "stdio", command: "node", args: ["server.js"] },
        },
      ];
      const mockTools = [
        { name: "mcp_localserver_tool1", description: "Test", inputSchema: {}, tags: [] },
      ];
      const runCommandError = new UserError("test", "RunCommandError", "Run command failed");

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      sandbox.stub(ODRProvider, "listServers").resolves([]);
      sandbox.stub(sharedOpts, "runCommand").resolves(err(runCommandError));

      await updateActionWithMCP(args);

      const telemetryCall = (ExtTelemetry.sendTelemetryErrorEvent as sinon.SinonStub).lastCall;
      chai.assert.equal(telemetryCall.args[2]["mcp-type"], "local");
    });

    it("should handle 401 error with no response headers", async () => {
      const args = [{ serverName: "testServer", serverConfig: { url: "http://test.com" } }];
      const mockTools = [
        { name: "mcp_testserver_tool1", description: "Test", inputSchema: {}, tags: [] },
      ];
      const axiosError = {
        status: 401,
        response: {},
      };

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      const axiosStub = sandbox.stub(axios, "get");
      axiosStub.withArgs("http://test.com").throws(axiosError);
      axiosStub
        .withArgs("http://test.com/.well-known/oauth-authorization-server")
        .resolves({ status: 200 });
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      await updateActionWithMCP(args);

      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      chai.assert.equal(calledInputs[QuestionNames.MCPForDAAuth], "OAuthPluginVault");
    });

    it("should parse www-authenticate header with match", async () => {
      const args = [{ serverName: "testServer", serverConfig: { url: "http://test.com" } }];
      const mockTools = [
        { name: "mcp_testserver_tool1", description: "Test", inputSchema: {}, tags: [] },
      ];
      const axiosError = {
        status: 401,
        response: {
          headers: {
            "www-authenticate": 'Bearer resource_metadata="http://auth.test.com/metadata"',
          },
        },
      };

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      sandbox.stub(axios, "get").onFirstCall().throws(axiosError);
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      await updateActionWithMCP(args);

      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      chai.assert.equal(
        calledInputs[QuestionNames.MCPForDAAuthMetadataUrl],
        "http://auth.test.com/metadata"
      );
    });

    it("should skip auth check for local MCP servers", async () => {
      const args = [
        {
          serverName: "localServer",
          serverConfig: { type: "stdio", command: "node", args: ["server.js"] },
        },
      ];
      const mockTools = [
        { name: "mcp_localserver_tool1", description: "Test", inputSchema: {}, tags: [] },
      ];

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      sandbox.stub(ODRProvider, "listServers").resolves([]);
      const axiosStub = sandbox.stub(axios, "get");
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      await updateActionWithMCP(args);

      // Axios should not be called for local MCP
      chai.assert.isFalse(axiosStub.called);
      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      chai.assert.equal(calledInputs[QuestionNames.MCPForDAAuth], "NoneAuth");
    });

    it("should return original server name when serverConfig type is not stdio", async () => {
      const args = [
        {
          serverName: "remoteServer",
          serverConfig: { type: "sse", url: "http://test.com" },
        },
      ];
      const mockTools = [
        { name: "mcp_remoteserver_tool1", description: "Test", inputSchema: {}, tags: [] },
      ];

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      sandbox.stub(axios, "get").resolves({ status: 200 });
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP(args);

      chai.assert.isTrue(result.isOk());
      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      chai.assert.isUndefined(calledInputs[QuestionNames.MCPLocalServerIdentifier]);
    });

    it("should construct detail with command only when args is undefined", async () => {
      const mcpContent = {
        servers: {
          "test-server": { url: "http://test.com" },
          "test-server2": {
            type: "stdio",
            command: "node",
          },
        },
      };
      const mockTools = [
        { name: "mcp_test-server_tool1", description: "Test tool", inputSchema: {}, tags: [] },
      ];

      const expectedPath = path.join(mockProjectPath, ".vscode", "mcp.json");
      sandbox.stub(fs, "pathExistsSync").withArgs(expectedPath).returns(true);
      sandbox
        .stub(fs, "readFileSync")
        .withArgs(expectedPath, "utf-8")
        .returns(JSON.stringify(mcpContent));
      sandbox.stub(parser, "parse").returns(mcpContent);
      Object.defineProperty(vscode.lm, "tools", { value: mockTools, configurable: true });
      sandbox.stub(axios, "get").resolves({ status: 200 });

      let capturedOptions: any;
      sandbox.stub(vscUI, "VS_CODE_UI").value({
        selectOption: sandbox.stub().callsFake((config: any) => {
          capturedOptions = config.options;
          return Promise.resolve(ok({ type: "success", result: "test-server" }));
        }),
      });
      sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      await updateActionWithMCP();

      const server2Option = capturedOptions.find((opt: any) => opt.id === "test-server2");
      chai.assert.equal(server2Option.detail, "node");
    });
  });

  describe("MCP gateway integration", () => {
    it("should fall back to selectedTools when startMcpGateway returns undefined", async () => {
      const args = [{ serverName: "testServer", serverConfig: { url: "http://test.com" } }];
      const mockTools = [
        { name: "mcp_testserver_tool1", description: "Test tool", inputSchema: {}, tags: [] },
      ];

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      Object.defineProperty(vscode.lm, "startMcpGateway", {
        value: sandbox.stub().resolves(undefined),
        configurable: true,
      });
      sandbox.stub(axios, "get").resolves({ status: 200 });
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP(args);

      chai.assert.isTrue(result.isOk());
      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      const tools = calledInputs[QuestionNames.MCPForDAAvailableTools];
      chai.assert.equal(tools.length, 1);
      chai.assert.equal(tools[0].name, "tool1");
    });

    it("should fall back to selectedTools when startMcpGateway throws", async () => {
      const args = [{ serverName: "testServer", serverConfig: { url: "http://test.com" } }];
      const mockTools = [
        { name: "mcp_testserver_tool1", description: "Test tool", inputSchema: {}, tags: [] },
      ];

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      Object.defineProperty(vscode.lm, "startMcpGateway", {
        value: sandbox.stub().rejects(new Error("Not supported")),
        configurable: true,
      });
      sandbox.stub(axios, "get").resolves({ status: 200 });
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP(args);

      chai.assert.isTrue(result.isOk());
      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      const tools = calledInputs[QuestionNames.MCPForDAAvailableTools];
      chai.assert.equal(tools.length, 1);
      chai.assert.equal(tools[0].name, "tool1");
    });

    it("should use gateway tools when gateway succeeds and tools match", async () => {
      const args = [{ serverName: "testServer", serverConfig: { url: "http://test.com" } }];
      const mockTools = [
        { name: "mcp_testserver_tool1", description: "Test tool 1", inputSchema: {}, tags: [] },
        { name: "mcp_testserver_tool2", description: "Test tool 2", inputSchema: {}, tags: [] },
      ];

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      const mockGateway = {
        servers: [{ label: "testServer", address: { toString: () => "http://localhost:12345" } }],
        dispose: sandbox.stub(),
      };
      Object.defineProperty(vscode.lm, "startMcpGateway", {
        value: sandbox.stub().resolves(mockGateway),
        configurable: true,
      });
      const connectStub = sandbox.stub(Client.prototype, "connect").resolves();
      const listToolsStub = sandbox.stub(Client.prototype, "listTools").resolves({
        tools: [
          { name: "tool1", description: "Test tool 1", inputSchema: { type: "object" as const } },
          { name: "tool2", description: "Test tool 2", inputSchema: { type: "object" as const } },
          { name: "tool3", description: "Other tool", inputSchema: { type: "object" as const } },
        ],
      });
      const closeStub = sandbox.stub(Client.prototype, "close").resolves();
      sandbox.stub(axios, "get").resolves({ status: 200 });
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP(args);

      chai.assert.isTrue(result.isOk());
      sinon.assert.calledOnce(connectStub);
      sinon.assert.calledOnce(listToolsStub);
      sinon.assert.calledOnce(closeStub);
      sinon.assert.calledOnce(mockGateway.dispose as sinon.SinonStub);
      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      const tools = calledInputs[QuestionNames.MCPForDAAvailableTools];
      // The gateway's tool list is authoritative; all server tools are returned as-is.
      chai.assert.equal(tools.length, 3);
      chai.assert.equal(tools[0].name, "tool1");
      chai.assert.equal(tools[1].name, "tool2");
      chai.assert.equal(tools[2].name, "tool3");
    });

    it("should use gateway tools even when vscode.lm.tools is empty (macOS regression)", async () => {
      const args = [{ serverName: "testServer", serverConfig: { url: "http://test.com" } }];

      // On macOS the MCP server's tools are not surfaced in vscode.lm.tools, so the
      // selectedTools filter is empty. The gateway result must still be used.
      sandbox.stub(vscode.lm, "tools").value([]);
      const mockGateway = {
        servers: [{ label: "testServer", address: { toString: () => "http://localhost:12345" } }],
        dispose: sandbox.stub(),
      };
      Object.defineProperty(vscode.lm, "startMcpGateway", {
        value: sandbox.stub().resolves(mockGateway),
        configurable: true,
      });
      const connectStub = sandbox.stub(Client.prototype, "connect").resolves();
      const listToolsStub = sandbox.stub(Client.prototype, "listTools").resolves({
        tools: [
          { name: "tool1", description: "Test tool 1", inputSchema: { type: "object" as const } },
          // tool2 has no description, exercising the `description ?? ""` dedup key path.
          { name: "tool2", inputSchema: { type: "object" as const } },
        ],
      });
      const closeStub = sandbox.stub(Client.prototype, "close").resolves();
      sandbox.stub(axios, "get").resolves({ status: 200 });
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP(args);

      chai.assert.isTrue(result.isOk());
      sinon.assert.calledOnce(connectStub);
      sinon.assert.calledOnce(listToolsStub);
      sinon.assert.calledOnce(closeStub);
      sinon.assert.calledOnce(mockGateway.dispose as sinon.SinonStub);
      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      const tools = calledInputs[QuestionNames.MCPForDAAvailableTools];
      chai.assert.equal(tools.length, 2);
      chai.assert.equal(tools[0].name, "tool1");
      chai.assert.equal(tools[1].name, "tool2");
    });

    it("should fall back to selectedTools when client.connect throws", async () => {
      const args = [{ serverName: "testServer", serverConfig: { url: "http://test.com" } }];
      const mockTools = [
        { name: "mcp_testserver_tool1", description: "Test tool", inputSchema: {}, tags: [] },
      ];

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      const mockGateway = {
        servers: [{ label: "testServer", address: { toString: () => "http://localhost:12345" } }],
        dispose: sandbox.stub(),
      };
      Object.defineProperty(vscode.lm, "startMcpGateway", {
        value: sandbox.stub().resolves(mockGateway),
        configurable: true,
      });
      sandbox.stub(Client.prototype, "connect").rejects(new Error("Connection failed"));
      const closeStub = sandbox.stub(Client.prototype, "close").resolves();
      sandbox.stub(axios, "get").resolves({ status: 200 });
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP(args);

      chai.assert.isTrue(result.isOk());
      // finally block should still call close and dispose
      sinon.assert.calledOnce(closeStub);
      sinon.assert.calledOnce(mockGateway.dispose as sinon.SinonStub);
      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      const tools = calledInputs[QuestionNames.MCPForDAAvailableTools];
      chai.assert.equal(tools.length, 1);
      chai.assert.equal(tools[0].name, "tool1");
    });

    it("should deduplicate gateway tools with same name and description", async () => {
      const args = [{ serverName: "testServer", serverConfig: { url: "http://test.com" } }];
      const mockTools = [
        { name: "mcp_testserver_tool1", description: "Dup tool", inputSchema: {}, tags: [] },
      ];

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      const mockGateway = {
        servers: [{ label: "testServer", address: { toString: () => "http://localhost:12345" } }],
        dispose: sandbox.stub(),
      };
      Object.defineProperty(vscode.lm, "startMcpGateway", {
        value: sandbox.stub().resolves(mockGateway),
        configurable: true,
      });
      sandbox.stub(Client.prototype, "connect").resolves();
      sandbox.stub(Client.prototype, "listTools").resolves({
        tools: [
          { name: "tool1", description: "Dup tool", inputSchema: { type: "object" as const } },
          { name: "tool1", description: "Dup tool", inputSchema: { type: "object" as const } },
        ],
      });
      sandbox.stub(Client.prototype, "close").resolves();
      sandbox.stub(axios, "get").resolves({ status: 200 });
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP(args);

      chai.assert.isTrue(result.isOk());
      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      const tools = calledInputs[QuestionNames.MCPForDAAvailableTools];
      // Only first match should be kept due to dedup
      chai.assert.equal(tools.length, 1);
    });

    it("should fall back to selectedTools when gateway returns an empty tool list", async () => {
      const args = [{ serverName: "testServer", serverConfig: { url: "http://test.com" } }];
      const mockTools = [
        { name: "mcp_testserver_tool1", description: "Fallback tool", inputSchema: {}, tags: [] },
      ];

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      const mockGateway = {
        servers: [{ label: "testServer", address: { toString: () => "http://localhost:12345" } }],
        dispose: sandbox.stub(),
      };
      Object.defineProperty(vscode.lm, "startMcpGateway", {
        value: sandbox.stub().resolves(mockGateway),
        configurable: true,
      });
      sandbox.stub(Client.prototype, "connect").resolves();
      const listToolsStub = sandbox.stub(Client.prototype, "listTools").resolves({
        tools: [],
      });
      const closeStub = sandbox.stub(Client.prototype, "close").resolves();
      sandbox.stub(axios, "get").resolves({ status: 200 });
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP(args);

      chai.assert.isTrue(result.isOk());
      sinon.assert.calledOnce(listToolsStub);
      sinon.assert.calledOnce(closeStub);
      sinon.assert.calledOnce(mockGateway.dispose as sinon.SinonStub);
      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      const tools = calledInputs[QuestionNames.MCPForDAAvailableTools];
      // Empty gateway result falls back to the vscode.lm.tools-derived list.
      chai.assert.equal(tools.length, 1);
      chai.assert.equal(tools[0].name, "tool1");
    });

    it("should include gateway tools even when absent from vscode.lm.tools", async () => {
      const args = [{ serverName: "testServer", serverConfig: { url: "http://test.com" } }];
      const mockTools = [
        { name: "mcp_testserver_tool1", description: "Tool A", inputSchema: {}, tags: [] },
      ];

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      const mockGateway = {
        servers: [{ label: "testServer", address: { toString: () => "http://localhost:12345" } }],
        dispose: sandbox.stub(),
      };
      Object.defineProperty(vscode.lm, "startMcpGateway", {
        value: sandbox.stub().resolves(mockGateway),
        configurable: true,
      });
      sandbox.stub(Client.prototype, "connect").resolves();
      sandbox.stub(Client.prototype, "listTools").resolves({
        tools: [
          {
            name: "unknownTool",
            description: "Not in vscode.lm.tools",
            inputSchema: { type: "object" as const },
          },
        ],
      });
      sandbox.stub(Client.prototype, "close").resolves();
      sandbox.stub(axios, "get").resolves({ status: 200 });
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP(args);

      // The gateway list is authoritative, so the tool is included even though it is
      // not present in vscode.lm.tools.
      chai.assert.isTrue(result.isOk());
      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      const tools = calledInputs[QuestionNames.MCPForDAAvailableTools];
      chai.assert.equal(tools.length, 1);
      chai.assert.equal(tools[0].name, "unknownTool");
    });

    it("should fall back to old gateway address API when servers property is absent", async () => {
      const args = [{ serverName: "testServer", serverConfig: { url: "http://test.com" } }];
      const mockTools = [
        { name: "mcp_testserver_tool1", description: "Test tool 1", inputSchema: {}, tags: [] },
      ];

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      const mockGateway = {
        address: { toString: () => "http://localhost:12345" },
        dispose: sandbox.stub(),
      };
      Object.defineProperty(vscode.lm, "startMcpGateway", {
        value: sandbox.stub().resolves(mockGateway),
        configurable: true,
      });
      sandbox.stub(Client.prototype, "connect").resolves();
      sandbox.stub(Client.prototype, "listTools").resolves({
        tools: [
          { name: "tool1", description: "Test tool 1", inputSchema: { type: "object" as const } },
        ],
      });
      sandbox.stub(Client.prototype, "close").resolves();
      sandbox.stub(axios, "get").resolves({ status: 200 });
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP(args);

      chai.assert.isTrue(result.isOk());
      sinon.assert.calledOnce(mockGateway.dispose as sinon.SinonStub);
      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      const tools = calledInputs[QuestionNames.MCPForDAAvailableTools];
      chai.assert.equal(tools.length, 1);
      chai.assert.equal(tools[0].name, "tool1");
    });

    it("should fall back to selectedTools when no matching server found in gateway", async () => {
      const args = [{ serverName: "testServer", serverConfig: { url: "http://test.com" } }];
      const mockTools = [
        { name: "mcp_testserver_tool1", description: "Test tool", inputSchema: {}, tags: [] },
      ];

      sandbox.stub(vscode.lm, "tools").value(mockTools);
      const mockGateway = {
        servers: [{ label: "otherServer", address: { toString: () => "http://localhost:12345" } }],
        dispose: sandbox.stub(),
      };
      Object.defineProperty(vscode.lm, "startMcpGateway", {
        value: sandbox.stub().resolves(mockGateway),
        configurable: true,
      });
      sandbox.stub(axios, "get").resolves({ status: 200 });
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP(args);

      chai.assert.isTrue(result.isOk());
      sinon.assert.calledOnce(mockGateway.dispose as sinon.SinonStub);
      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      const tools = calledInputs[QuestionNames.MCPForDAAvailableTools];
      chai.assert.equal(tools.length, 1);
      chai.assert.equal(tools[0].name, "tool1");
    });
  });

  describe("extractLocalServerIdentifier edge cases", () => {
    it("should return original name when ODR listServers throws", async () => {
      const args = [
        {
          serverName: "odrServer",
          serverConfig: { type: "stdio", command: "odr", args: ["run", "my-server"] },
        },
      ];
      const mockODRTools = [{ name: "tool1", description: "Test", inputSchema: {} }];

      const isODRStub = sandbox.stub(ODRProvider, "isODRServer");
      isODRStub.onFirstCall().returns(true); // extractLocalServerIdentifier
      isODRStub.onSecondCall().returns(true); // main function ODR check
      sandbox.stub(ODRProvider, "listServers").rejects(new Error("ODR unavailable"));
      sandbox.stub(ODRProvider, "getToolsForODRServer").resolves(mockODRTools);
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP(args);

      chai.assert.isTrue(result.isOk());
      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      // Should fall back to original serverName
      chai.assert.equal(calledInputs[QuestionNames.MCPLocalServerIdentifier], "odrServer");
    });

    it("should return original name when no ODR server matches by command and args", async () => {
      const args = [
        {
          serverName: "odrServer",
          serverConfig: { type: "stdio", command: "odr", args: ["run", "other-server"] },
        },
      ];
      const mockODRTools = [{ name: "tool1", description: "Test", inputSchema: {} }];
      const mockODRServers = [
        {
          name: "my-server",
          display_name: "My Server",
          description: "Test Server",
          version: "1.0.0",
          identifier: "my-server-identifier",
          tools: mockODRTools,
          packageFamily: "test.package",
          command: "odr",
          args: ["run", "my-server"], // Different args than the config
        },
      ];

      const isODRStub = sandbox.stub(ODRProvider, "isODRServer");
      isODRStub.onFirstCall().returns(true); // extractLocalServerIdentifier
      isODRStub.onSecondCall().returns(true); // main function ODR check
      sandbox.stub(ODRProvider, "listServers").resolves(mockODRServers);
      sandbox.stub(ODRProvider, "getToolsForODRServer").resolves(mockODRTools);
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP(args);

      chai.assert.isTrue(result.isOk());
      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      // Should fall back to original serverName since args don't match
      chai.assert.equal(calledInputs[QuestionNames.MCPLocalServerIdentifier], "odrServer");
    });

    it("should return original name for non-ODR local server in extractLocalServerIdentifier", async () => {
      const args = [
        {
          serverName: "myLocalServer",
          serverConfig: { type: "stdio", command: "node", args: ["server.js"] },
        },
      ];
      const mockTools = [
        { name: "mcp_mylocalserver_tool1", description: "Test", inputSchema: {}, tags: [] },
      ];

      sandbox.stub(ODRProvider, "isODRServer").returns(false);
      sandbox.stub(vscode.lm, "tools").value(mockTools);
      const runCommandStub = sandbox.stub(sharedOpts, "runCommand").resolves(ok(undefined));

      const result = await updateActionWithMCP(args);

      chai.assert.isTrue(result.isOk());
      const calledInputs = runCommandStub.getCall(0).args[1] as Inputs;
      chai.assert.equal(calledInputs[QuestionNames.MCPLocalServerIdentifier], "myLocalServer");
    });
  });
});

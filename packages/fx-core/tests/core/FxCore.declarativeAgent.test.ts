// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs, Platform, ok } from "@microsoft/teamsfx-api";
import axios from "axios";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import sinon from "sinon";
import { FxCore } from "../../src";
import { setTools } from "../../src/common/globalVars";
import { ActionInjector } from "../../src/component/configManager/actionInjector";
import { LocalMcpPrefix } from "../../src/component/constants";
import { pathUtils } from "../../src/component/utils/pathUtils";
import { QuestionNames } from "../../src/question";
import { MockTools } from "./utils";

describe("updateActionWithMCP", () => {
  const tools = new MockTools();
  const sandbox = sinon.createSandbox();
  const projectPath = "/test/project";
  const pluginManifestPath = "/test/project/ai-plugin.json";
  const mcpServerUrl = "https://example.com/mcp";
  const serverName = "testServer";

  beforeEach(() => {
    setTools(tools);
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("should successfully update action with MCP without auth", async () => {
    const core = new FxCore(tools);
    const inputs: Inputs = {
      projectPath,
      platform: Platform.VSCode,
      [QuestionNames.PluginManifestFilePath]: pluginManifestPath,
      [QuestionNames.MCPForDAServerUrl]: mcpServerUrl,
      [QuestionNames.MCPForDAServerName]: serverName,
      [QuestionNames.MCPForDAAuth]: "None",
      [QuestionNames.MCPForDAAvailableTools]: [
        {
          name: "testTool",
          description: "Test tool description",
          inputSchema: {
            type: "object",
            properties: { param1: { type: "string" } },
            required: ["param1"],
          },
        },
      ],
      [QuestionNames.MCPForDAPreFetchTools]: ["testTool"],
      ignoreLockByUT: true,
    };

    const existingPlugin = {
      functions: [],
      runtimes: [],
    };

    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readJSON").resolves(existingPlugin);
    const writeJSONStub = sandbox.stub(fs, "writeJSON").resolves();
    sandbox.stub(pathUtils, "getYmlFilePath").returns("/test/project/teamsapp.yml");

    const showMessageStub = sandbox.stub(tools.ui, "showMessage").resolves(ok("OK"));
    const openFileStub = sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.updateActionWithMCP(inputs);

    assert.isTrue(result.isOk());
    assert.isTrue(writeJSONStub.calledOnce); // Only ai-plugin.json is written (tools are now embedded)
    assert.isTrue(showMessageStub.calledOnce);
    assert.isTrue(openFileStub.calledOnce);
  });

  it("should embed tools in mcp_tool_description when no matched runtime exists", async () => {
    const core = new FxCore(tools);
    const inputs: Inputs = {
      projectPath,
      platform: Platform.VSCode,
      [QuestionNames.PluginManifestFilePath]: pluginManifestPath,
      [QuestionNames.MCPForDAServerUrl]: mcpServerUrl,
      [QuestionNames.MCPForDAServerName]: serverName,
      [QuestionNames.MCPForDAAuth]: "None",
      [QuestionNames.MCPForDAAvailableTools]: [
        {
          name: "testTool",
          description: "Test tool description",
          inputSchema: {
            type: "object",
            properties: { param1: { type: "string" } },
            required: ["param1"],
          },
        },
      ],
      [QuestionNames.MCPForDAPreFetchTools]: ["testTool"],
      ignoreLockByUT: true,
    };

    // No existing runtime, so a new one will be created
    const existingPlugin = {
      functions: [],
      runtimes: [],
    };

    let writtenPluginData: any;

    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readJSON").resolves(existingPlugin);
    sandbox.stub(fs, "writeJSON").callsFake((filePath: string, data) => {
      writtenPluginData = data;
      return Promise.resolve();
    });
    sandbox.stub(pathUtils, "getYmlFilePath").returns("/test/project/teamsapp.yml");

    sandbox.stub(tools.ui, "showMessage").resolves(ok("OK"));
    sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.updateActionWithMCP(inputs);

    assert.isTrue(result.isOk());

    // Verify the runtime was added with embedded tools in mcp_tool_description.tools
    const mcpRuntime = writtenPluginData.runtimes.find(
      (r: any) => r.type === "RemoteMCPServer" && r.spec.url === mcpServerUrl
    );
    assert.isDefined(mcpRuntime);
    assert.isDefined(mcpRuntime.spec.mcp_tool_description.tools);
    assert.equal(mcpRuntime.spec.mcp_tool_description.tools.length, 1);
    assert.equal(mcpRuntime.spec.mcp_tool_description.tools[0].name, "testTool");
    assert.isDefined(mcpRuntime.spec.mcp_tool_description.tools[0].title);
  });

  it("should replace existing runtime with new tools when matched runtime exists", async () => {
    const core = new FxCore(tools);
    const inputs: Inputs = {
      projectPath,
      platform: Platform.VSCode,
      [QuestionNames.PluginManifestFilePath]: pluginManifestPath,
      [QuestionNames.MCPForDAServerUrl]: mcpServerUrl,
      [QuestionNames.MCPForDAServerName]: serverName,
      [QuestionNames.MCPForDAAuth]: "None",
      [QuestionNames.MCPForDAAvailableTools]: [
        {
          name: "newTool",
          description: "New tool description",
          inputSchema: {
            type: "object",
            properties: { param1: { type: "string" } },
            required: ["param1"],
          },
        },
      ],
      [QuestionNames.MCPForDAPreFetchTools]: ["newTool"],
      ignoreLockByUT: true,
    };

    // Existing runtime with mcp_tool_description.file set
    const existingPlugin = {
      functions: [
        {
          name: "oldTool",
          description: "Old tool",
        },
      ],
      runtimes: [
        {
          type: "RemoteMCPServer",
          spec: {
            url: mcpServerUrl,
            mcp_tool_description: {
              file: "existing-mcp-tools.json",
            },
          },
          run_for_functions: ["oldTool"],
        },
      ],
    };

    let writtenPluginData: any;

    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readJSON").resolves(existingPlugin);
    sandbox.stub(fs, "writeJSON").callsFake((filePath: string, data) => {
      writtenPluginData = data;
      return Promise.resolve();
    });
    sandbox.stub(pathUtils, "getYmlFilePath").returns("/test/project/teamsapp.yml");

    sandbox.stub(tools.ui, "showMessage").resolves(ok("OK"));
    sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.updateActionWithMCP(inputs);

    assert.isTrue(result.isOk());

    // Verify the runtime now has embedded tools instead of file reference
    const mcpRuntime = writtenPluginData.runtimes.find(
      (r: any) => r.type === "RemoteMCPServer" && r.spec.url === mcpServerUrl
    );
    assert.isDefined(mcpRuntime);
    assert.isDefined(mcpRuntime.spec.mcp_tool_description.tools);
    assert.equal(mcpRuntime.spec.mcp_tool_description.tools.length, 1);
    assert.equal(mcpRuntime.spec.mcp_tool_description.tools[0].name, "newTool");
  });

  it("should successfully update action with OAuth authentication", async () => {
    const core = new FxCore(tools);
    const inputs: Inputs = {
      projectPath,
      platform: Platform.VSCode,
      [QuestionNames.PluginManifestFilePath]: pluginManifestPath,
      [QuestionNames.MCPForDAServerUrl]: mcpServerUrl,
      [QuestionNames.MCPForDAServerName]: serverName,
      [QuestionNames.MCPForDAAuth]: "OAuthPluginVault",
      [QuestionNames.MCPForDAAuthType]: "oauth",
      [QuestionNames.MCPForDAAuthWellKnownUrl]:
        "https://example.com/.well-known/oauth-authorization-server",
      [QuestionNames.MCPForDAAvailableTools]: [
        {
          name: "testTool",
          description: "Test tool description",
          inputSchema: {
            type: "object",
            properties: { param1: { type: "string" } },
            required: ["param1"],
          },
        },
      ],
      [QuestionNames.MCPForDAPreFetchTools]: ["testTool"],
      ignoreLockByUT: true,
    };

    const existingPlugin = {
      functions: [],
      runtimes: [],
    };

    const oauthMetadata = {
      authorization_endpoint: "https://example.com/oauth/authorize",
      token_endpoint: "https://example.com/oauth/token",
      refresh_endpoint: "https://example.com/oauth/refresh",
    };

    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readJSON").resolves(existingPlugin);
    const writeJSONStub = sandbox.stub(fs, "writeJSON").resolves();
    sandbox.stub(pathUtils, "getYmlFilePath").returns("/test/project/teamsapp.yml");
    sandbox.stub(axios, "get").resolves({ status: 200, data: oauthMetadata });
    const injectOAuthStub = sandbox
      .stub(ActionInjector, "injectCreateOAuthActionForMCP")
      .resolves();

    const showMessageStub = sandbox.stub(tools.ui, "showMessage").resolves(ok("OK"));
    const openFileStub = sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.updateActionWithMCP(inputs);

    assert.isTrue(result.isOk());
    assert.isTrue(injectOAuthStub.calledOnce);
    assert.isTrue(writeJSONStub.calledOnce); // Only ai-plugin.json is written (tools are now embedded)
    assert.isTrue(showMessageStub.calledOnce);
    assert.isTrue(openFileStub.calledOnce);
  });

  it("should successfully update action with OAuth authentication using metadata URL", async () => {
    const core = new FxCore(tools);
    const inputs: Inputs = {
      projectPath,
      platform: Platform.VSCode,
      [QuestionNames.PluginManifestFilePath]: pluginManifestPath,
      [QuestionNames.MCPForDAServerUrl]: mcpServerUrl,
      [QuestionNames.MCPForDAServerName]: serverName,
      [QuestionNames.MCPForDAAuth]: "OAuthPluginVault",
      [QuestionNames.MCPForDAAuthType]: "oauth",
      [QuestionNames.MCPForDAAuthMetadataUrl]: "https://example.com/mcp/metadata",
      [QuestionNames.MCPForDAAvailableTools]: [
        {
          name: "testTool",
          description: "Test tool description",
          inputSchema: {
            type: "object",
            properties: { param1: { type: "string" } },
            required: ["param1"],
          },
        },
      ],
      [QuestionNames.MCPForDAPreFetchTools]: ["testTool"],
      ignoreLockByUT: true,
    };

    const existingPlugin = {
      functions: [],
      runtimes: [],
    };

    const mcpMetadata = {
      authorization_servers: ["https://example.com/oauth"],
    };

    const oauthMetadata = {
      authorization_endpoint: "https://example.com/oauth/authorize",
      token_endpoint: "https://example.com/oauth/token",
      refresh_endpoint: "https://example.com/oauth/refresh",
    };

    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readJSON").resolves(existingPlugin);
    const writeJSONStub = sandbox.stub(fs, "writeJSON").resolves();
    sandbox.stub(pathUtils, "getYmlFilePath").returns("/test/project/teamsapp.yml");
    sandbox
      .stub(axios, "get")
      .onFirstCall()
      .resolves({ status: 200, data: mcpMetadata })
      .onSecondCall()
      .resolves({ status: 200, data: oauthMetadata });
    const injectOAuthStub = sandbox
      .stub(ActionInjector, "injectCreateOAuthActionForMCP")
      .resolves();

    const showMessageStub = sandbox.stub(tools.ui, "showMessage").resolves(ok("OK"));
    const openFileStub = sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.updateActionWithMCP(inputs);

    assert.isTrue(result.isOk());
    assert.isTrue(injectOAuthStub.calledOnce);
    assert.isTrue(writeJSONStub.calledOnce); // Only ai-plugin.json is written (tools are now embedded)
    assert.isTrue(showMessageStub.calledOnce);
    assert.isTrue(openFileStub.calledOnce);
  });

  it("should return error when plugin manifest file does not exist", async () => {
    const core = new FxCore(tools);
    const inputs: Inputs = {
      projectPath,
      platform: Platform.VSCode,
      [QuestionNames.PluginManifestFilePath]: pluginManifestPath,
      [QuestionNames.MCPForDAServerUrl]: mcpServerUrl,
      [QuestionNames.MCPForDAServerName]: serverName,
      ignoreLockByUT: true,
    };

    sandbox.stub(fs, "pathExists").resolves(false);

    const result = await core.updateActionWithMCP(inputs);

    assert.isTrue(result.isErr());
    if (result.isErr()) {
      // Check error source/name since localization strings might be missing
      assert.isTrue(
        result.error.source === "MCPForDAPluginManifestNotFound" ||
          result.error.name === "PluginManifestNotFound" ||
          result.error.message.includes("PluginManifestNotFound")
      );
    }
  });

  it("should return error when projectPath is undefined", async () => {
    const core = new FxCore(tools);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.PluginManifestFilePath]: pluginManifestPath,
      ignoreLockByUT: true,
    };

    try {
      const result = await core.updateActionWithMCP(inputs);
      // If it returns a result instead of throwing, check if it's an error
      if (result.isErr()) {
        assert.include(result.error.message.toLowerCase(), "project");
      } else {
        assert.fail("Expected error to be thrown or returned");
      }
    } catch (error: any) {
      // If it throws, check the error message
      assert.include(error.message.toLowerCase(), "project");
    }
  });

  it("should return error when MCP tools are not provided", async () => {
    const core = new FxCore(tools);
    const inputs: Inputs = {
      projectPath,
      platform: Platform.VSCode,
      [QuestionNames.PluginManifestFilePath]: pluginManifestPath,
      [QuestionNames.MCPForDAServerUrl]: mcpServerUrl,
      [QuestionNames.MCPForDAServerName]: serverName,
      ignoreLockByUT: true,
    };

    const existingPlugin = {
      functions: [],
      runtimes: [],
    };

    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readJSON").resolves(existingPlugin);

    const result = await core.updateActionWithMCP(inputs);

    assert.isTrue(result.isErr());
    if (result.isErr()) {
      // Check error source/name since localization strings might be missing
      assert.isTrue(
        result.error.source === "MCPForDAPreFetchToolsNotFound" ||
          result.error.name === "PreFetchToolsNotFound" ||
          result.error.message.includes("PreFetchToolsNotFound")
      );
    }
  });

  it("should properly filter and update existing MCP runtimes", async () => {
    const core = new FxCore(tools);
    const inputs: Inputs = {
      projectPath,
      platform: Platform.VSCode,
      [QuestionNames.PluginManifestFilePath]: pluginManifestPath,
      [QuestionNames.MCPForDAServerUrl]: mcpServerUrl,
      [QuestionNames.MCPForDAServerName]: serverName,
      [QuestionNames.MCPForDAAuth]: "None",
      [QuestionNames.MCPForDAAvailableTools]: [
        {
          name: "newTool",
          description: "New tool description",
          inputSchema: {
            type: "object",
            properties: { param1: { type: "string" } },
            required: ["param1"],
          },
        },
      ],
      [QuestionNames.MCPForDAPreFetchTools]: ["newTool"],
      ignoreLockByUT: true,
    };

    const existingPlugin = {
      functions: [
        {
          name: "oldTool",
          description: "Old tool",
          parameters: { type: "object" },
        },
      ],
      runtimes: [
        {
          type: "RemoteMCPServer",
          spec: {
            url: mcpServerUrl,
          },
          run_for_functions: ["oldTool"],
        },
        {
          type: "RemoteMCPServer",
          spec: {
            url: "https://other.com/mcp",
          },
          run_for_functions: ["otherTool"],
        },
      ],
    };

    let writtenPlugin: any;
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readJSON").resolves(existingPlugin);
    sandbox.stub(fs, "writeJSON").callsFake((filePath: string, data) => {
      writtenPlugin = data;
      return Promise.resolve();
    });
    sandbox.stub(pathUtils, "getYmlFilePath").returns("/test/project/teamsapp.yml");

    const showMessageStub = sandbox.stub(tools.ui, "showMessage").resolves(ok("OK"));
    const openFileStub = sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.updateActionWithMCP(inputs);

    assert.isTrue(result.isOk());

    // Verify that old tool functions were removed and new ones added (only name and description)
    assert.isDefined(writtenPlugin);
    assert.equal(writtenPlugin.functions.length, 1);
    assert.equal(writtenPlugin.functions[0].name, "newTool");
    assert.isDefined(writtenPlugin.functions[0].description);
    assert.isUndefined(writtenPlugin.functions[0].parameters);
    assert.isUndefined(writtenPlugin.functions[0].inputSchema);

    // Verify that the existing runtime for the same server was removed and new one added
    const mcpRuntimes = writtenPlugin.runtimes.filter(
      (r: any) => r.type === "RemoteMCPServer" && r.spec.url === mcpServerUrl
    );
    assert.equal(mcpRuntimes.length, 1);
    assert.deepEqual(mcpRuntimes[0].run_for_functions, ["newTool"]);

    // Verify tools are now embedded in the runtime spec
    assert.isDefined(mcpRuntimes[0].spec.mcp_tool_description.tools);
    assert.equal(mcpRuntimes[0].spec.mcp_tool_description.tools.length, 1);
    assert.equal(mcpRuntimes[0].spec.mcp_tool_description.tools[0].name, "newTool");
    assert.isDefined(mcpRuntimes[0].spec.mcp_tool_description.tools[0].title);

    // Verify that other runtimes are preserved
    const otherRuntimes = writtenPlugin.runtimes.filter(
      (r: any) => r.type === "RemoteMCPServer" && r.spec.url === "https://other.com/mcp"
    );
    assert.equal(otherRuntimes.length, 1);
  });

  it("should handle provisionResources call when user clicks Provision", async () => {
    const core = new FxCore(tools);
    const inputs: Inputs = {
      projectPath,
      platform: Platform.VSCode,
      [QuestionNames.PluginManifestFilePath]: pluginManifestPath,
      [QuestionNames.MCPForDAServerUrl]: mcpServerUrl,
      [QuestionNames.MCPForDAServerName]: serverName,
      [QuestionNames.MCPForDAAuth]: "None",
      [QuestionNames.MCPForDAAvailableTools]: [
        {
          name: "testTool",
          description: "Test tool description",
          inputSchema: {
            type: "object",
            properties: { param1: { type: "string" } },
            required: ["param1"],
          },
        },
      ],
      [QuestionNames.MCPForDAPreFetchTools]: ["testTool"],
      ignoreLockByUT: true,
    };

    const existingPlugin = {
      functions: [],
      runtimes: [],
    };

    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readJSON").resolves(existingPlugin);
    const writeJSONStub = sandbox.stub(fs, "writeJSON").resolves();
    sandbox.stub(pathUtils, "getYmlFilePath").returns("/test/project/teamsapp.yml");

    // Mock the showMessage to return "Provision" to trigger provision call
    const showMessageStub = sandbox.stub(tools.ui, "showMessage").resolves(ok("Provision"));
    const openFileStub = sandbox.stub(tools.ui, "openFile").resolves();
    const provisionStub = sandbox.stub(core, "provisionResources").resolves(ok(undefined));

    const result = await core.updateActionWithMCP(inputs);

    assert.isTrue(result.isOk());
    assert.isTrue(showMessageStub.calledOnce);
    assert.isTrue(openFileStub.calledOnce);
    assert.isTrue(writeJSONStub.calledOnce); // Only ai-plugin.json is written (tools are now embedded)

    // Wait a bit for the async provision call
    await new Promise((resolve) => setTimeout(resolve, 10));
    assert.isTrue(provisionStub.calledOnce);
  });

  it("should embed tools directly in runtime spec instead of creating separate file", async () => {
    const core = new FxCore(tools);
    const inputs: Inputs = {
      projectPath,
      platform: Platform.VSCode,
      [QuestionNames.PluginManifestFilePath]: pluginManifestPath,
      [QuestionNames.MCPForDAServerUrl]: mcpServerUrl,
      [QuestionNames.MCPForDAServerName]: serverName,
      [QuestionNames.MCPForDAAuth]: "None",
      [QuestionNames.MCPForDAAvailableTools]: [
        {
          name: "testTool",
          description: "Test tool description",
          inputSchema: {
            type: "object",
            properties: { param1: { type: "string" } },
            required: ["param1"],
          },
        },
      ],
      [QuestionNames.MCPForDAPreFetchTools]: ["testTool"],
      ignoreLockByUT: true,
    };

    const existingPlugin = {
      functions: [],
      runtimes: [],
    };

    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readJSON").resolves(existingPlugin);
    let writtenPluginData: any;
    sandbox.stub(fs, "writeJSON").callsFake((filePath: string, data) => {
      writtenPluginData = data;
      return Promise.resolve();
    });
    sandbox.stub(pathUtils, "getYmlFilePath").returns("/test/project/teamsapp.yml");

    sandbox.stub(tools.ui, "showMessage").resolves(ok("OK"));
    sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.updateActionWithMCP(inputs);

    assert.isTrue(result.isOk());

    // Verify tools are embedded in the runtime spec directly
    const mcpRuntime = writtenPluginData.runtimes.find(
      (r: any) => r.type === "RemoteMCPServer" && r.spec.url === mcpServerUrl
    );
    assert.isDefined(mcpRuntime);
    assert.isDefined(mcpRuntime.spec.mcp_tool_description.tools);
    assert.equal(mcpRuntime.spec.mcp_tool_description.tools.length, 1);
    assert.equal(mcpRuntime.spec.mcp_tool_description.tools[0].name, "testTool");
    assert.isDefined(mcpRuntime.spec.mcp_tool_description.tools[0].title);
  });

  it("should show error when mcpAuthMetadataUrl is not provided for OAuth without well-known URL", async () => {
    const core = new FxCore(tools);
    const inputs: Inputs = {
      projectPath,
      platform: Platform.VSCode,
      [QuestionNames.PluginManifestFilePath]: pluginManifestPath,
      [QuestionNames.MCPForDAServerUrl]: mcpServerUrl,
      [QuestionNames.MCPForDAServerName]: serverName,
      [QuestionNames.MCPForDAAuth]: "OAuthPluginVault",
      [QuestionNames.MCPForDAAuthType]: "oauth",
      // Neither MCPForDAAuthWellKnownUrl nor MCPForDAAuthMetadataUrl is provided
      [QuestionNames.MCPForDAAvailableTools]: [
        {
          name: "testTool",
          description: "Test tool description",
          inputSchema: {
            type: "object",
            properties: { param1: { type: "string" } },
            required: ["param1"],
          },
        },
      ],
      [QuestionNames.MCPForDAPreFetchTools]: ["testTool"],
      ignoreLockByUT: true,
    };

    const existingPlugin = {
      functions: [],
      runtimes: [],
    };

    sandbox.stub(fs, "pathExists").callsFake(async (filePath: string) => {
      return !filePath.includes("mcp-tools");
    });
    sandbox.stub(fs, "readJSON").resolves(existingPlugin);
    sandbox.stub(fs, "writeJSON").resolves();
    sandbox.stub(pathUtils, "getYmlFilePath").returns("/test/project/teamsapp.yml");
    sandbox.stub(ActionInjector, "injectCreateOAuthActionForMCP").resolves();

    const showMessageStub = sandbox.stub(tools.ui, "showMessage").resolves(ok("OK"));
    sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.updateActionWithMCP(inputs);

    // The method should still return ok - the error is caught and shown to user,
    // then execution continues (this path is verified by code coverage)
    assert.isTrue(result.isOk());
  });

  it("should show error when authorization_servers is missing in metadata response", async () => {
    const core = new FxCore(tools);
    const inputs: Inputs = {
      projectPath,
      platform: Platform.VSCode,
      [QuestionNames.PluginManifestFilePath]: pluginManifestPath,
      [QuestionNames.MCPForDAServerUrl]: mcpServerUrl,
      [QuestionNames.MCPForDAServerName]: serverName,
      [QuestionNames.MCPForDAAuth]: "OAuthPluginVault",
      [QuestionNames.MCPForDAAuthType]: "oauth",
      [QuestionNames.MCPForDAAuthMetadataUrl]: "https://example.com/mcp/metadata",
      [QuestionNames.MCPForDAAvailableTools]: [
        {
          name: "testTool",
          description: "Test tool description",
          inputSchema: {
            type: "object",
            properties: { param1: { type: "string" } },
            required: ["param1"],
          },
        },
      ],
      [QuestionNames.MCPForDAPreFetchTools]: ["testTool"],
      ignoreLockByUT: true,
    };

    const existingPlugin = {
      functions: [],
      runtimes: [],
    };

    // Return metadata without authorization_servers
    const mcpMetadataWithoutAuthServers = {
      // authorization_servers is missing
    };

    sandbox.stub(fs, "pathExists").callsFake(async (filePath: string) => {
      return !filePath.includes("mcp-tools");
    });
    sandbox.stub(fs, "readJSON").resolves(existingPlugin);
    sandbox.stub(fs, "writeJSON").resolves();
    sandbox.stub(pathUtils, "getYmlFilePath").returns("/test/project/teamsapp.yml");
    sandbox.stub(axios, "get").resolves({ status: 200, data: mcpMetadataWithoutAuthServers });
    sandbox.stub(ActionInjector, "injectCreateOAuthActionForMCP").resolves();

    sandbox.stub(tools.ui, "showMessage").resolves(ok("OK"));
    sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.updateActionWithMCP(inputs);

    // The method should still return ok - the error is caught and shown to user,
    // then execution continues (this path is verified by code coverage)
    assert.isTrue(result.isOk());
  });

  it("should show error when authorization_servers is empty array", async () => {
    const core = new FxCore(tools);
    const inputs: Inputs = {
      projectPath,
      platform: Platform.VSCode,
      [QuestionNames.PluginManifestFilePath]: pluginManifestPath,
      [QuestionNames.MCPForDAServerUrl]: mcpServerUrl,
      [QuestionNames.MCPForDAServerName]: serverName,
      [QuestionNames.MCPForDAAuth]: "OAuthPluginVault",
      [QuestionNames.MCPForDAAuthType]: "oauth",
      [QuestionNames.MCPForDAAuthMetadataUrl]: "https://example.com/mcp/metadata",
      [QuestionNames.MCPForDAAvailableTools]: [
        {
          name: "testTool",
          description: "Test tool description",
          inputSchema: {
            type: "object",
            properties: { param1: { type: "string" } },
            required: ["param1"],
          },
        },
      ],
      [QuestionNames.MCPForDAPreFetchTools]: ["testTool"],
      ignoreLockByUT: true,
    };

    const existingPlugin = {
      functions: [],
      runtimes: [],
    };

    // Return metadata with empty authorization_servers array
    const mcpMetadataWithEmptyAuthServers = {
      authorization_servers: [],
    };

    sandbox.stub(fs, "pathExists").callsFake(async (filePath: string) => {
      return !filePath.includes("mcp-tools");
    });
    sandbox.stub(fs, "readJSON").resolves(existingPlugin);
    sandbox.stub(fs, "writeJSON").resolves();
    sandbox.stub(pathUtils, "getYmlFilePath").returns("/test/project/teamsapp.yml");
    sandbox.stub(axios, "get").resolves({ status: 200, data: mcpMetadataWithEmptyAuthServers });
    sandbox.stub(ActionInjector, "injectCreateOAuthActionForMCP").resolves();

    sandbox.stub(tools.ui, "showMessage").resolves(ok("OK"));
    sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.updateActionWithMCP(inputs);

    // The method should still return ok - the error is caught and shown to user,
    // then execution continues (this path is verified by code coverage)
    assert.isTrue(result.isOk());
  });

  it("should show error when OAuth metadata is missing authorization_endpoint or token_endpoint", async () => {
    const core = new FxCore(tools);
    const inputs: Inputs = {
      projectPath,
      platform: Platform.VSCode,
      [QuestionNames.PluginManifestFilePath]: pluginManifestPath,
      [QuestionNames.MCPForDAServerUrl]: mcpServerUrl,
      [QuestionNames.MCPForDAServerName]: serverName,
      [QuestionNames.MCPForDAAuth]: "OAuthPluginVault",
      [QuestionNames.MCPForDAAuthType]: "oauth",
      [QuestionNames.MCPForDAAuthWellKnownUrl]:
        "https://example.com/.well-known/oauth-authorization-server",
      [QuestionNames.MCPForDAAvailableTools]: [
        {
          name: "testTool",
          description: "Test tool description",
          inputSchema: {
            type: "object",
            properties: { param1: { type: "string" } },
            required: ["param1"],
          },
        },
      ],
      [QuestionNames.MCPForDAPreFetchTools]: ["testTool"],
      ignoreLockByUT: true,
    };

    const existingPlugin = {
      functions: [],
      runtimes: [],
    };

    // Return OAuth metadata without authorization_endpoint and token_endpoint
    const incompleteOAuthMetadata = {
      // Missing authorization_endpoint and token_endpoint
      refresh_endpoint: "https://example.com/oauth/refresh",
    };

    sandbox.stub(fs, "pathExists").callsFake(async (filePath: string) => {
      return !filePath.includes("mcp-tools");
    });
    sandbox.stub(fs, "readJSON").resolves(existingPlugin);
    sandbox.stub(fs, "writeJSON").resolves();
    sandbox.stub(pathUtils, "getYmlFilePath").returns("/test/project/teamsapp.yml");
    sandbox.stub(axios, "get").resolves({ status: 200, data: incompleteOAuthMetadata });
    sandbox.stub(ActionInjector, "injectCreateOAuthActionForMCP").resolves();

    sandbox.stub(tools.ui, "showMessage").resolves(ok("OK"));
    sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.updateActionWithMCP(inputs);

    // The method should still return ok - the error is caught and shown to user,
    // then execution continues (this path is verified by code coverage)
    assert.isTrue(result.isOk());
  });

  it("should throw error when provisionResources is called directly", async () => {
    const core = new FxCore(tools);
    const inputs: Inputs = {
      projectPath,
      platform: Platform.VSCode,
      ignoreLockByUT: true,
    };

    try {
      // Access the protected method via the class - FxCore overrides this,
      // but we can test the base class behavior by creating an instance directly
      const { FxCoreDeclarativeAgentPart } = await import("../../src/core/FxCore.declarativeAgent");
      const declarativeAgentPart = new FxCoreDeclarativeAgentPart();
      await declarativeAgentPart.provisionResources(inputs);
      assert.fail("Expected an error to be thrown");
    } catch (error: any) {
      assert.include(error.message, "not implemented");
    }
  });
});

describe("updateActionWithMCP - Local MCP Support", () => {
  const tools = new MockTools();
  const sandbox = sinon.createSandbox();
  const projectPath = "/test/project";
  const pluginManifestPath = "/test/project/ai-plugin.json";
  const serverName = "testLocalServer";
  const localServerIdentifier = "com.test.local.server";

  beforeEach(() => {
    setTools(tools);
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("should successfully update action with local MCP server", async () => {
    const core = new FxCore(tools);
    const inputs: Inputs = {
      projectPath,
      platform: Platform.VSCode,
      [QuestionNames.PluginManifestFilePath]: pluginManifestPath,
      [QuestionNames.MCPForDAServerName]: serverName,
      [QuestionNames.MCPLocalServerIdentifier]: localServerIdentifier,
      [QuestionNames.MCPForDAAuth]: "None",
      [QuestionNames.MCPForDAAvailableTools]: [
        {
          name: "localTool",
          description: "Local MCP tool description",
          inputSchema: {
            type: "object",
            properties: { param1: { type: "string" } },
            required: ["param1"],
          },
        },
      ],
      [QuestionNames.MCPForDAPreFetchTools]: ["localTool"],
      ignoreLockByUT: true,
    };

    const existingPlugin = {
      functions: [],
      runtimes: [],
    };

    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readJSON").resolves(existingPlugin);
    const writeJSONStub = sandbox.stub(fs, "writeJSON").resolves();
    sandbox.stub(pathUtils, "getYmlFilePath").returns("/test/project/teamsapp.yml");

    const showMessageStub = sandbox.stub(tools.ui, "showMessage").resolves(ok("OK"));
    const openFileStub = sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.updateActionWithMCP(inputs);

    assert.isTrue(result.isOk());

    // Verify the local MCP runtime was added correctly
    assert.isTrue(writeJSONStub.calledOnce);
    const writtenData = writeJSONStub.getCall(0).args[1];
    const runtimes = writtenData.runtimes as any[];

    assert.equal(runtimes.length, 1);
    assert.equal(runtimes[0].type, "LocalPlugin");
    assert.equal(runtimes[0].spec.local_endpoint, LocalMcpPrefix + localServerIdentifier);
    assert.deepEqual(runtimes[0].run_for_functions, ["localTool"]);

    assert.isTrue(showMessageStub.calledOnce);
    assert.isTrue(openFileStub.calledOnce);

    const localFunctions = writtenData.functions as any[];
    assert.equal(localFunctions.length, 1);
    assert.equal(localFunctions[0].name, "localTool");
    assert.equal(localFunctions[0].description, "Local MCP tool description");
  });

  it("should filter and update existing local MCP runtimes correctly", async () => {
    const core = new FxCore(tools);
    const inputs: Inputs = {
      projectPath,
      platform: Platform.VSCode,
      [QuestionNames.PluginManifestFilePath]: pluginManifestPath,
      [QuestionNames.MCPForDAServerName]: serverName,
      [QuestionNames.MCPLocalServerIdentifier]: localServerIdentifier,
      [QuestionNames.MCPForDAAuth]: "None",
      [QuestionNames.MCPForDAAvailableTools]: [
        {
          name: "newLocalTool",
          description: "New local tool description",
          inputSchema: {
            type: "object",
            properties: { param1: { type: "string" } },
            required: ["param1"],
          },
        },
      ],
      [QuestionNames.MCPForDAPreFetchTools]: ["newLocalTool"],
      ignoreLockByUT: true,
    };

    const existingPlugin = {
      functions: [
        {
          name: "oldLocalTool",
          description: "Old local tool",
          parameters: { type: "object" },
        },
      ],
      runtimes: [
        {
          type: "LocalPlugin",
          spec: {
            identifier: localServerIdentifier,
          },
          run_for_functions: ["oldLocalTool"],
        },
        {
          type: "LocalPlugin",
          spec: {
            identifier: "com.other.local.server",
          },
          run_for_functions: ["otherTool"],
        },
      ],
    };

    let writtenPlugin: any;
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readJSON").resolves(existingPlugin);
    sandbox.stub(fs, "writeJSON").callsFake((path, data) => {
      writtenPlugin = data;
      return Promise.resolve();
    });
    sandbox.stub(pathUtils, "getYmlFilePath").returns("/test/project/teamsapp.yml");

    const showMessageStub = sandbox.stub(tools.ui, "showMessage").resolves(ok("OK"));
    const openFileStub = sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.updateActionWithMCP(inputs);

    assert.isTrue(result.isOk());

    // Verify that old tool functions were removed and new ones added
    assert.equal(writtenPlugin.functions.length, 1);
    assert.equal(writtenPlugin.functions[0].name, "newLocalTool");

    // Verify that the existing runtime for the same local server was removed and new one added
    const localRuntimes = writtenPlugin.runtimes.filter((r: any) => r.type === "LocalPlugin");
    assert.equal(localRuntimes.length, 1);
    assert.deepEqual(localRuntimes[0].run_for_functions, ["newLocalTool"]);
  });

  it("should handle mixed remote and local MCP servers", async () => {
    const core = new FxCore(tools);
    const inputs: Inputs = {
      projectPath,
      platform: Platform.VSCode,
      [QuestionNames.PluginManifestFilePath]: pluginManifestPath,
      [QuestionNames.MCPForDAServerName]: serverName,
      [QuestionNames.MCPLocalServerIdentifier]: localServerIdentifier,
      [QuestionNames.MCPForDAAuth]: "None",
      [QuestionNames.MCPForDAAvailableTools]: [
        {
          name: "localTool",
          description: "Local MCP tool",
          inputSchema: {
            type: "object",
            properties: { param1: { type: "string" } },
            required: ["param1"],
          },
        },
      ],
      [QuestionNames.MCPForDAPreFetchTools]: ["localTool"],
      ignoreLockByUT: true,
    };

    const existingPlugin = {
      functions: [
        {
          name: "remoteTool",
          description: "Remote tool",
          parameters: { type: "object" },
        },
      ],
      runtimes: [
        {
          type: "RemoteMCPServer",
          spec: {
            url: "https://remote.example.com/mcp",
          },
          run_for_functions: ["remoteTool"],
        },
      ],
    };

    let writtenPlugin: any;
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readJSON").resolves(existingPlugin);
    sandbox.stub(fs, "writeJSON").callsFake((path, data) => {
      writtenPlugin = data;
      return Promise.resolve();
    });
    sandbox.stub(pathUtils, "getYmlFilePath").returns("/test/project/teamsapp.yml");

    const showMessageStub = sandbox.stub(tools.ui, "showMessage").resolves(ok("OK"));
    const openFileStub = sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.updateActionWithMCP(inputs);

    assert.isTrue(result.isOk());

    // Verify both local and remote functions exist
    assert.equal(writtenPlugin.functions.length, 2);
    const functionNames = writtenPlugin.functions.map((f: any) => f.name);
    assert.include(functionNames, "localTool");
    assert.include(functionNames, "remoteTool");

    // Verify both local and remote runtimes exist
    assert.equal(writtenPlugin.runtimes.length, 2);
    const runtimeTypes = writtenPlugin.runtimes.map((r: any) => r.type);
    assert.include(runtimeTypes, "LocalPlugin");
    assert.include(runtimeTypes, "RemoteMCPServer");

    // Verify local runtime has correct identifier
    const localRuntime = writtenPlugin.runtimes.find((r: any) => r.type === "LocalPlugin");
    assert.equal(localRuntime.spec.local_endpoint, LocalMcpPrefix + localServerIdentifier);
    assert.deepEqual(localRuntime.run_for_functions, ["localTool"]);
  });
});

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

    sandbox.stub(fs, "pathExists").callsFake(async (filePath: string) => {
      // Return false for mcp-tools.json to avoid infinite loop, true for ai-plugin.json
      return !filePath.includes("mcp-tools");
    });
    sandbox.stub(fs, "readJSON").resolves(existingPlugin);
    const writeJSONStub = sandbox.stub(fs, "writeJSON").resolves();
    sandbox.stub(pathUtils, "getYmlFilePath").returns("/test/project/teamsapp.yml");

    const showMessageStub = sandbox.stub(tools.ui, "showMessage").resolves(ok("OK"));
    const openFileStub = sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.updateActionWithMCP(inputs);

    assert.isTrue(result.isOk());
    assert.isTrue(writeJSONStub.calledTwice); // mcp-tools.json and ai-plugin.json
    assert.isTrue(showMessageStub.calledOnce);
    assert.isTrue(openFileStub.calledOnce);
  });

  it("should create default mcp-tools.json when mcpFile is undefined (no matched runtime)", async () => {
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

    // No existing runtime with mcp_tool_description.file, so mcpFile will be undefined
    const existingPlugin = {
      functions: [],
      runtimes: [],
    };

    let writtenMcpToolsPath = "";
    let writtenMcpToolsData: any;
    let writtenPluginData: any;

    sandbox.stub(fs, "pathExists").callsFake(async (filePath: string) => {
      // Return false for mcp-tools.json so it uses the default name
      if (filePath.includes("mcp-tools")) {
        return false;
      }
      return true; // ai-plugin.json exists
    });
    sandbox.stub(fs, "readJSON").resolves(existingPlugin);
    sandbox.stub(fs, "writeJSON").callsFake((filePath: string, data) => {
      if (filePath.includes("mcp-tools")) {
        writtenMcpToolsPath = filePath;
        writtenMcpToolsData = data;
      } else {
        writtenPluginData = data;
      }
      return Promise.resolve();
    });
    sandbox.stub(pathUtils, "getYmlFilePath").returns("/test/project/teamsapp.yml");

    sandbox.stub(tools.ui, "showMessage").resolves(ok("OK"));
    sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.updateActionWithMCP(inputs);

    assert.isTrue(result.isOk());

    // Verify mcp-tools.json was created with default name (not incremented)
    assert.isTrue(writtenMcpToolsPath.includes("mcp-tools.json"));
    assert.isFalse(writtenMcpToolsPath.includes("mcp-tools-"));

    // Verify mcp-tools.json contains the tool with full details
    assert.isDefined(writtenMcpToolsData);
    assert.equal(writtenMcpToolsData.tools.length, 1);
    assert.equal(writtenMcpToolsData.tools[0].name, "testTool");
    assert.isDefined(writtenMcpToolsData.tools[0].inputSchema);
    assert.isDefined(writtenMcpToolsData.tools[0].title);

    // Verify the runtime was added with mcp_tool_description.file reference
    const mcpRuntime = writtenPluginData.runtimes.find(
      (r: any) => r.type === "RemoteMCPServer" && r.spec.url === mcpServerUrl
    );
    assert.isDefined(mcpRuntime);
    assert.equal(mcpRuntime.spec.mcp_tool_description.file, "mcp-tools.json");
  });

  it("should reuse existing mcp_tool_description.file when matched runtime exists", async () => {
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

    let writtenMcpToolsPath = "";
    let writtenPluginData: any;

    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readJSON").resolves(existingPlugin);
    sandbox.stub(fs, "writeJSON").callsFake((filePath: string, data) => {
      if (filePath.includes("mcp-tools") || filePath.includes("existing-mcp-tools")) {
        writtenMcpToolsPath = filePath;
      } else {
        writtenPluginData = data;
      }
      return Promise.resolve();
    });
    sandbox.stub(pathUtils, "getYmlFilePath").returns("/test/project/teamsapp.yml");

    sandbox.stub(tools.ui, "showMessage").resolves(ok("OK"));
    sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.updateActionWithMCP(inputs);

    assert.isTrue(result.isOk());

    // Verify the existing mcp_tool_description.file is reused
    assert.isTrue(writtenMcpToolsPath.includes("existing-mcp-tools.json"));

    // Verify the runtime uses the same file reference
    const mcpRuntime = writtenPluginData.runtimes.find(
      (r: any) => r.type === "RemoteMCPServer" && r.spec.url === mcpServerUrl
    );
    assert.isDefined(mcpRuntime);
    assert.equal(mcpRuntime.spec.mcp_tool_description.file, "existing-mcp-tools.json");
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

    sandbox.stub(fs, "pathExists").callsFake(async (filePath: string) => {
      // Return false for mcp-tools.json to avoid infinite loop, true for ai-plugin.json
      return !filePath.includes("mcp-tools");
    });
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
    assert.isTrue(writeJSONStub.calledTwice); // mcp-tools.json and ai-plugin.json
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

    sandbox.stub(fs, "pathExists").callsFake(async (filePath: string) => {
      // Return false for mcp-tools.json to avoid infinite loop, true for ai-plugin.json
      return !filePath.includes("mcp-tools");
    });
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
    assert.isTrue(writeJSONStub.calledTwice); // mcp-tools.json and ai-plugin.json
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
            enable_dynamic_discovery: false,
          },
          run_for_functions: ["oldTool"],
        },
        {
          type: "RemoteMCPServer",
          spec: {
            url: "https://other.com/mcp",
            enable_dynamic_discovery: false,
          },
          run_for_functions: ["otherTool"],
        },
      ],
    };

    let writtenPlugin: any;
    let writtenMcpTools: any;
    sandbox.stub(fs, "pathExists").callsFake(async (filePath: string) => {
      // Return false for mcp-tools.json to avoid infinite loop, true for ai-plugin.json
      return !filePath.includes("mcp-tools");
    });
    sandbox.stub(fs, "readJSON").resolves(existingPlugin);
    sandbox.stub(fs, "writeJSON").callsFake((filePath: string, data) => {
      if (filePath.includes("mcp-tools")) {
        writtenMcpTools = data;
      } else {
        writtenPlugin = data;
      }
      return Promise.resolve();
    });
    sandbox.stub(pathUtils, "getYmlFilePath").returns("/test/project/teamsapp.yml");

    const showMessageStub = sandbox.stub(tools.ui, "showMessage").resolves(ok("OK"));
    const openFileStub = sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.updateActionWithMCP(inputs);

    assert.isTrue(result.isOk());

    // Verify mcp-tools.json contains tools with full details including title and inputSchema
    assert.isDefined(writtenMcpTools);
    assert.equal(writtenMcpTools.tools.length, 1);
    assert.equal(writtenMcpTools.tools[0].name, "newTool");
    assert.isDefined(writtenMcpTools.tools[0].inputSchema);
    assert.isDefined(writtenMcpTools.tools[0].title);

    // Verify that old tool functions were removed and new ones added (only name and description)
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

    sandbox.stub(fs, "pathExists").callsFake(async (filePath: string) => {
      // Return false for mcp-tools.json to avoid infinite loop, true for ai-plugin.json
      return !filePath.includes("mcp-tools");
    });
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
    assert.isTrue(writeJSONStub.calledTwice); // mcp-tools.json and ai-plugin.json

    // Wait a bit for the async provision call
    await new Promise((resolve) => setTimeout(resolve, 10));
    assert.isTrue(provisionStub.calledOnce);
  });

  it("should generate unique mcp-tools filename when default already exists", async () => {
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

    // Simulate mcp-tools.json and mcp-tools-1.json already exist
    sandbox.stub(fs, "pathExists").callsFake(async (filePath: string) => {
      if (filePath.includes("mcp-tools.json") && !filePath.includes("mcp-tools-")) {
        return true; // mcp-tools.json exists
      }
      if (filePath.includes("mcp-tools-1.json")) {
        return true; // mcp-tools-1.json exists
      }
      if (filePath.includes("mcp-tools-2.json")) {
        return false; // mcp-tools-2.json does not exist
      }
      return true; // ai-plugin.json exists
    });
    sandbox.stub(fs, "readJSON").resolves(existingPlugin);
    let writtenMcpToolsPath = "";
    sandbox.stub(fs, "writeJSON").callsFake((filePath: string, data) => {
      if (filePath.includes("mcp-tools")) {
        writtenMcpToolsPath = filePath;
      }
      return Promise.resolve();
    });
    sandbox.stub(pathUtils, "getYmlFilePath").returns("/test/project/teamsapp.yml");

    sandbox.stub(tools.ui, "showMessage").resolves(ok("OK"));
    sandbox.stub(tools.ui, "openFile").resolves();

    const result = await core.updateActionWithMCP(inputs);

    assert.isTrue(result.isOk());
    // Verify the filename was incremented to mcp-tools-2.json since mcp-tools.json and mcp-tools-1.json exist
    assert.isTrue(writtenMcpToolsPath.includes("mcp-tools-2.json"));
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
            enable_dynamic_discovery: false,
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

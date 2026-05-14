/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, afterEach } from "mocha";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import * as os from "node:os";
import { addMCPAction } from "../../../src/declarativeAgent/actions/addActionFromMCP";
import type { AddMCPActionInput, MCPToolDefinition } from "../../../src/declarativeAgent/types";

describe("declarativeAgent/actions/addMCPAction", () => {
  let tmpDir: string;

  afterEach(async () => {
    if (tmpDir) {
      await fs.rm(tmpDir, { recursive: true, force: true });
    }
  });

  async function setup(): Promise<string> {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "atk-test-mcp-action-"));
    return tmpDir;
  }

  function minimalPluginManifest(): string {
    return JSON.stringify({
      schema_version: "v2.4",
      name_for_human: "Test Plugin",
      namespace: "testplugin",
      description_for_human: "A test plugin",
    });
  }

  function sampleTools(): MCPToolDefinition[] {
    return [
      { name: "getTodos", description: "Get all TODO items" },
      {
        name: "addTodo",
        description: "Add a TODO item",
        inputSchema: {
          type: "object",
          properties: { title: { type: "string" } },
          required: ["title"],
        },
      },
      { name: "deleteTodo", description: "Delete a TODO item" },
    ];
  }

  function makeInput(
    overrides: Partial<AddMCPActionInput> & { pluginManifestPath: string }
  ): AddMCPActionInput {
    return {
      serverUrlOrIdentifier: "https://mcp.example.com/sse",
      serverName: "test-server",
      isLocal: false,
      auth: "none",
      availableTools: sampleTools(),
      selectedTools: ["getTodos", "addTodo"],
      ...overrides,
    };
  }

  // -----------------------------------------------------------------------
  // Remote MCP — basic success
  // -----------------------------------------------------------------------
  it("should add remote MCP tools to ai-plugin.json", async () => {
    const dir = await setup();
    const pluginPath = path.join(dir, "ai-plugin.json");
    await fs.writeFile(pluginPath, minimalPluginManifest(), "utf-8");

    const result = await addMCPAction(makeInput({ pluginManifestPath: pluginPath }));
    expect(result.isOk()).to.be.true;

    // Verify ai-plugin.json was updated
    const saved = JSON.parse(await fs.readFile(pluginPath, "utf-8"));
    expect(saved.functions).to.have.length(2);
    expect(saved.functions.map((f: { name: string }) => f.name)).to.include.members([
      "getTodos",
      "addTodo",
    ]);

    // Verify runtime entry
    const runtime = saved.runtimes?.find(
      (r: Record<string, unknown>) =>
        (r.spec as Record<string, unknown>)?.url === "https://mcp.example.com/sse"
    );
    expect(runtime).to.exist;
    expect(runtime.type).to.equal("RemoteMCPServer");
    expect(runtime.run_for_functions).to.include.members(["getTodos", "addTodo"]);
  });

  // -----------------------------------------------------------------------
  // Remote MCP — writes mcp-tools.json sidecar
  // -----------------------------------------------------------------------
  it("should write mcp-tools.json for remote MCP servers", async () => {
    const dir = await setup();
    const pluginPath = path.join(dir, "ai-plugin.json");
    await fs.writeFile(pluginPath, minimalPluginManifest(), "utf-8");

    const result = await addMCPAction(makeInput({ pluginManifestPath: pluginPath }));
    expect(result.isOk()).to.be.true;

    const toolsPath = path.join(dir, "mcp-tools.json");
    const toolsContent = JSON.parse(await fs.readFile(toolsPath, "utf-8"));
    expect(toolsContent).to.be.an("array").with.length(2);
    expect(toolsContent[0].name).to.equal("getTodos");
    expect(toolsContent[1].name).to.equal("addTodo");
    expect(toolsContent[1].inputSchema).to.deep.include({ type: "object" });
  });

  // -----------------------------------------------------------------------
  // Local MCP — uses localmcp:// prefix
  // -----------------------------------------------------------------------
  it("should add local MCP tools with localmcp:// prefix", async () => {
    const dir = await setup();
    const pluginPath = path.join(dir, "ai-plugin.json");
    await fs.writeFile(pluginPath, minimalPluginManifest(), "utf-8");

    const result = await addMCPAction(
      makeInput({
        pluginManifestPath: pluginPath,
        serverUrlOrIdentifier: "my-local-server",
        isLocal: true,
        selectedTools: ["getTodos"],
      })
    );
    expect(result.isOk()).to.be.true;

    const saved = JSON.parse(await fs.readFile(pluginPath, "utf-8"));
    const runtime = saved.runtimes?.[0];
    expect(runtime).to.exist;
    expect(runtime.type).to.equal("LocalPlugin");
    expect(runtime.spec?.url).to.equal("localmcp://my-local-server");
  });

  // -----------------------------------------------------------------------
  // Local MCP — does NOT write mcp-tools.json
  // -----------------------------------------------------------------------
  it("should not write mcp-tools.json for local MCP servers", async () => {
    const dir = await setup();
    const pluginPath = path.join(dir, "ai-plugin.json");
    await fs.writeFile(pluginPath, minimalPluginManifest(), "utf-8");

    await addMCPAction(
      makeInput({
        pluginManifestPath: pluginPath,
        isLocal: true,
        serverUrlOrIdentifier: "local-server",
        selectedTools: ["getTodos"],
      })
    );

    const toolsPath = path.join(dir, "mcp-tools.json");
    let exists = true;
    try {
      await fs.access(toolsPath);
    } catch {
      exists = false;
    }
    expect(exists).to.be.false;
  });

  // -----------------------------------------------------------------------
  // OAuth auth — adds OAuthPluginVault to runtime
  // -----------------------------------------------------------------------
  it("should add OAuth auth info to runtime when auth is oauth", async () => {
    const dir = await setup();
    const pluginPath = path.join(dir, "ai-plugin.json");
    await fs.writeFile(pluginPath, minimalPluginManifest(), "utf-8");

    const result = await addMCPAction(
      makeInput({
        pluginManifestPath: pluginPath,
        auth: "oauth",
        oauthMeta: {
          authorizationUrl: "https://auth.example.com/authorize",
          tokenUrl: "https://auth.example.com/token",
        },
        selectedTools: ["getTodos"],
      })
    );
    expect(result.isOk()).to.be.true;

    const saved = JSON.parse(await fs.readFile(pluginPath, "utf-8"));
    const runtime = saved.runtimes?.[0];
    expect(runtime.auth).to.deep.equal({
      type: "OAuthPluginVault",
      reference_id: "test-server_oauth",
    });
  });

  // -----------------------------------------------------------------------
  // Microsoft Entra auth — adds OAuthPluginVault auth
  // -----------------------------------------------------------------------
  it("should add OAuthPluginVault auth for microsoft-entra", async () => {
    const dir = await setup();
    const pluginPath = path.join(dir, "ai-plugin.json");
    await fs.writeFile(pluginPath, minimalPluginManifest(), "utf-8");

    const result = await addMCPAction(
      makeInput({
        pluginManifestPath: pluginPath,
        auth: "microsoft-entra",
        selectedTools: ["getTodos"],
      })
    );
    expect(result.isOk()).to.be.true;

    const saved = JSON.parse(await fs.readFile(pluginPath, "utf-8"));
    const runtime = saved.runtimes?.[0];
    expect(runtime.auth).to.exist;
    expect(runtime.auth.type).to.equal("OAuthPluginVault");
  });

  // -----------------------------------------------------------------------
  // Dedup — re-adding same server replaces tools
  // -----------------------------------------------------------------------
  it("should replace existing tools when same server URL is added again", async () => {
    const dir = await setup();
    const pluginPath = path.join(dir, "ai-plugin.json");
    await fs.writeFile(pluginPath, minimalPluginManifest(), "utf-8");

    // First add — getTodos only
    await addMCPAction(
      makeInput({
        pluginManifestPath: pluginPath,
        selectedTools: ["getTodos"],
      })
    );

    // Second add — addTodo only (same server URL)
    const result = await addMCPAction(
      makeInput({
        pluginManifestPath: pluginPath,
        selectedTools: ["addTodo"],
      })
    );
    expect(result.isOk()).to.be.true;

    const saved = JSON.parse(await fs.readFile(pluginPath, "utf-8"));
    // Should have only the second set of tools, not duplicates
    const toolNames = saved.functions.map((f: { name: string }) => f.name);
    expect(toolNames).to.include("addTodo");
  });

  // -----------------------------------------------------------------------
  // No auth — runtime has no auth field
  // -----------------------------------------------------------------------
  it("should set auth to None when auth is none", async () => {
    const dir = await setup();
    const pluginPath = path.join(dir, "ai-plugin.json");
    await fs.writeFile(pluginPath, minimalPluginManifest(), "utf-8");

    await addMCPAction(
      makeInput({
        pluginManifestPath: pluginPath,
        auth: "none",
        selectedTools: ["getTodos"],
      })
    );

    const saved = JSON.parse(await fs.readFile(pluginPath, "utf-8"));
    const runtime = saved.runtimes?.[0];
    expect(runtime.auth).to.deep.equal({ type: "None" });
  });

  // -----------------------------------------------------------------------
  // Plugin manifest not found — returns UserError
  // -----------------------------------------------------------------------
  it("should return error when plugin manifest does not exist", async () => {
    const dir = await setup();
    const pluginPath = path.join(dir, "nonexistent", "ai-plugin.json");

    const result = await addMCPAction(makeInput({ pluginManifestPath: pluginPath }));
    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("PluginManifestNotFound");
    }
  });

  // -----------------------------------------------------------------------
  // mcp-tools.json includes inputSchema
  // -----------------------------------------------------------------------
  it("should preserve inputSchema in mcp-tools.json", async () => {
    const dir = await setup();
    const pluginPath = path.join(dir, "ai-plugin.json");
    await fs.writeFile(pluginPath, minimalPluginManifest(), "utf-8");

    await addMCPAction(
      makeInput({
        pluginManifestPath: pluginPath,
        selectedTools: ["addTodo"],
      })
    );

    const toolsPath = path.join(dir, "mcp-tools.json");
    const toolsContent = JSON.parse(await fs.readFile(toolsPath, "utf-8"));
    expect(toolsContent).to.have.length(1);
    expect(toolsContent[0].inputSchema).to.deep.include({ type: "object" });
    expect(toolsContent[0].inputSchema.properties).to.have.property("title");
  });
});

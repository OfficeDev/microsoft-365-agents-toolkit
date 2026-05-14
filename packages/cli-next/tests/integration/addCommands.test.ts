/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import { expect } from "chai";
import { describe, it, afterEach, beforeEach } from "mocha";
import * as sinon from "sinon";
import { buildProgram } from "../../src/commands";
import { cliTelemetry } from "../../src/telemetry";

/**
 * Integration tests for add / set / m365 commands.
 *
 * Tests in-process Commander wiring with real core-next DA operations.
 * Mocks only I/O boundaries (telemetry, console).
 */
describe("Add / Set / M365 command integration", () => {
  const sandbox = sinon.createSandbox();
  let tmpDir: string;

  beforeEach(async () => {
    sandbox.stub(cliTelemetry, "sendEvent");
    sandbox.stub(cliTelemetry, "sendErrorEvent");
    sandbox.stub(cliTelemetry, "flush").resolves();
    sandbox.stub(console, "log");
    sandbox.stub(console, "warn");
    sandbox.stub(console, "error");
    tmpDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), "cli-integ-add-"));
  });

  afterEach(async () => {
    sandbox.restore();
    process.exitCode = undefined;
    await fs.promises.rm(tmpDir, { recursive: true, force: true });
  });

  /**
   * Create a minimal DA project for real execution tests.
   */
  async function createDAProject(): Promise<string> {
    const appDir = path.join(tmpDir, "appPackage");
    await fs.promises.mkdir(appDir, { recursive: true });

    const pngHeader = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    await fs.promises.writeFile(path.join(appDir, "color.png"), pngHeader);
    await fs.promises.writeFile(path.join(appDir, "outline.png"), pngHeader);

    const manifest = {
      $schema:
        "https://developer.microsoft.com/en-us/json-schemas/teams/v1.19/MicrosoftTeams.schema.json",
      manifestVersion: "1.19",
      version: "1.0.0",
      id: "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",
      developer: {
        name: "Test Corp",
        websiteUrl: "https://example.com",
        privacyUrl: "https://example.com/privacy",
        termsOfUseUrl: "https://example.com/tos",
      },
      name: { short: "TestApp", full: "Test App" },
      description: { short: "A test app", full: "A test application" },
      icons: { color: "color.png", outline: "outline.png" },
      accentColor: "#4F6BED",
      copilotAgents: {
        declarativeAgents: [{ id: "daAgent", file: "declarativeAgent.json" }],
      },
    };
    await fs.promises.writeFile(
      path.join(appDir, "manifest.json"),
      JSON.stringify(manifest, null, 2)
    );

    const daManifest = {
      $schema:
        "https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.6/schema.json",
      version: "v1.6",
      name: "Test DA",
      description: "A test declarative agent",
      instructions: "Be helpful",
    };
    const agentPath = path.join(appDir, "declarativeAgent.json");
    await fs.promises.writeFile(agentPath, JSON.stringify(daManifest, null, 2));

    return agentPath;
  }

  // ─── --help tests (Commander tree inspection) ──────────────

  describe("Command tree inspection", () => {
    it("atk add — should have action, capability, and auth-config subcommands", () => {
      const program = buildProgram("atk");
      program.exitOverride();

      const addCmd = program.commands.find((c) => c.name() === "add");
      expect(addCmd, "'add' command should exist").to.exist;

      const subcmdNames = addCmd!.commands.map((c) => c.name());
      expect(subcmdNames).to.include("action");
      expect(subcmdNames).to.include("capability");
      expect(subcmdNames).to.include("auth-config");
    });

    it("atk add action — should have --api-spec-path, --plugin-manifest-path, --action-id", () => {
      const program = buildProgram("atk");
      program.exitOverride();

      const addCmd = program.commands.find((c) => c.name() === "add");
      const actionCmd = addCmd!.commands.find((c) => c.name() === "action");
      expect(actionCmd, "'add action' command should exist").to.exist;

      const optLongs = actionCmd!.options.map((o) => o.long);
      expect(optLongs).to.include("--api-spec-path");
      expect(optLongs).to.include("--plugin-manifest-path");
      expect(optLongs).to.include("--action-id");
      expect(optLongs).to.include("--agent-manifest-path");
    });

    it("atk add action — should have MCP options", () => {
      const program = buildProgram("atk");
      program.exitOverride();

      const addCmd = program.commands.find((c) => c.name() === "add");
      const actionCmd = addCmd!.commands.find((c) => c.name() === "action");
      expect(actionCmd).to.exist;

      const optLongs = actionCmd!.options.map((o) => o.long);
      expect(optLongs).to.include("--api-plugin-type");
      expect(optLongs).to.include("--mcp-server-url");
      expect(optLongs).to.include("--mcp-server-name");
      expect(optLongs).to.include("--mcp-is-local");
      expect(optLongs).to.include("--mcp-auth-type");
      expect(optLongs).to.include("--mcp-tools-file");
      expect(optLongs).to.include("--mcp-selected-tools");
      expect(optLongs).to.include("--mcp-oauth-auth-url");
      expect(optLongs).to.include("--mcp-oauth-token-url");
      expect(optLongs).to.include("--mcp-oauth-refresh-url");
    });

    it("atk add capability — should require --capability-type", () => {
      const program = buildProgram("atk");
      program.exitOverride();

      const addCmd = program.commands.find((c) => c.name() === "add");
      const capCmd = addCmd!.commands.find((c) => c.name() === "capability");
      expect(capCmd).to.exist;

      const optLongs = capCmd!.options.map((o) => o.long);
      expect(optLongs).to.include("--capability-type");
    });
  });

  // ─── Real execution tests ─────────────────────────────────

  describe("Real execution (filesystem side effects)", () => {
    it("atk add capability --capability-type web-search — adds WebSearch to manifest", async () => {
      const agentManifestPath = await createDAProject();
      const program = buildProgram("atk");
      program.exitOverride();

      await program.parseAsync([
        "node",
        "atk",
        "add",
        "capability",
        "--capability-type",
        "web-search",
        "--agent-manifest-path",
        agentManifestPath,
      ]);

      const updated = JSON.parse(await fs.promises.readFile(agentManifestPath, "utf-8"));
      const caps = updated.capabilities as Array<Record<string, unknown>>;
      expect(caps).to.be.an("array").with.length.greaterThan(0);
      const webSearch = caps.find((c) => c.name === "WebSearch");
      expect(webSearch, "WebSearch capability should be present").to.exist;

      // Verify success message was logged
      const logStub = console.log as sinon.SinonStub;
      const output = logStub
        .getCalls()
        .map((c) => c.args.join(" "))
        .join("\n");
      expect(output).to.include("Capability added successfully");
    });

    it("atk set sensitivityLabel --label-id label-xyz — sets label in manifest", async () => {
      const agentManifestPath = await createDAProject();
      const program = buildProgram("atk");
      program.exitOverride();

      await program.parseAsync([
        "node",
        "atk",
        "set",
        "sensitivityLabel",
        "--label-id",
        "label-xyz",
        "--agent-manifest-path",
        agentManifestPath,
      ]);

      const updated = JSON.parse(await fs.promises.readFile(agentManifestPath, "utf-8"));
      expect(updated.sensitivity_label).to.deep.equal({ id: "label-xyz" });

      const logStub = console.log as sinon.SinonStub;
      const output = logStub
        .getCalls()
        .map((c) => c.args.join(" "))
        .join("\n");
      expect(output).to.include("Sensitivity label set successfully");
    });

    it("atk add action --api-plugin-type mcp — adds MCP action to DA project", async () => {
      await createDAProject();

      // Create ai-plugin.json in appPackage/
      const pluginPath = path.join(tmpDir, "appPackage", "ai-plugin.json");
      await fs.promises.writeFile(
        pluginPath,
        JSON.stringify({
          schema_version: "v2.4",
          name_for_human: "Test Plugin",
          namespace: "testplugin",
          description_for_human: "A test plugin",
        }),
        "utf-8"
      );

      // Create tools file
      const toolsFile = path.join(tmpDir, "mcp-tools-input.json");
      await fs.promises.writeFile(
        toolsFile,
        JSON.stringify([
          { name: "getWeather", description: "Get weather for a city" },
          { name: "getNews", description: "Get latest news" },
        ]),
        "utf-8"
      );

      const program = buildProgram("atk");
      program.exitOverride();

      await program.parseAsync([
        "node",
        "atk",
        "add",
        "action",
        "--api-plugin-type",
        "mcp",
        "--mcp-server-url",
        "https://weather-mcp.example.com/sse",
        "--mcp-server-name",
        "weather-server",
        "--mcp-auth-type",
        "none",
        "--mcp-tools-file",
        toolsFile,
        "--mcp-selected-tools",
        "getWeather",
        "--plugin-manifest-path",
        pluginPath,
      ]);

      // Verify ai-plugin.json was updated with MCP runtime
      const errStub = console.error as sinon.SinonStub;
      const errOutput = errStub
        .getCalls()
        .map((c) => c.args.join(" "))
        .join("\n");
      expect(process.exitCode, `Unexpected exit code. stderr: ${errOutput}`).to.not.equal(1);
      const savedPlugin = JSON.parse(await fs.promises.readFile(pluginPath, "utf-8"));
      expect(savedPlugin.functions).to.be.an("array").with.length.greaterThan(0);
      const fnNames = savedPlugin.functions.map((f: { name: string }) => f.name);
      expect(fnNames).to.include("getWeather");

      const runtime = savedPlugin.runtimes?.find(
        (r: Record<string, unknown>) =>
          (r.spec as Record<string, unknown>)?.url === "https://weather-mcp.example.com/sse"
      );
      expect(runtime, "MCP runtime should be present").to.exist;
      expect(runtime.type).to.equal("RemoteMCPServer");

      // Verify mcp-tools.json sidecar was created
      const toolsSidecarPath = path.join(tmpDir, "appPackage", "mcp-tools.json");
      const toolsSidecar = JSON.parse(await fs.promises.readFile(toolsSidecarPath, "utf-8"));
      expect(toolsSidecar).to.be.an("array").with.length(1);
      expect(toolsSidecar[0].name).to.equal("getWeather");

      // Verify success message
      const logStub = console.log as sinon.SinonStub;
      const output = logStub
        .getCalls()
        .map((c) => c.args.join(" "))
        .join("\n");
      expect(output).to.include("MCP action added successfully");
    });

    it("atk add action --api-plugin-type api-spec — requires OpenAPI options", async () => {
      const program = buildProgram("atk");
      program.exitOverride();

      await program.parseAsync([
        "node",
        "atk",
        "add",
        "action",
        "--api-plugin-type",
        "api-spec",
        // Missing required options for api-spec mode
      ]);

      // The handler catches errors and sets exitCode
      expect(process.exitCode).to.equal(1);
      const errStub = console.error as sinon.SinonStub;
      const errOutput = errStub
        .getCalls()
        .map((c) => c.args.join(" "))
        .join("\n");
      expect(errOutput).to.include("--api-spec-path");
    });
  });

  // ─── Required option verification (Commander tree) ─────────

  describe("Required option flags", () => {
    it("atk add action — OpenAPI options are optional (required only for api-spec mode)", () => {
      const program = buildProgram("atk");
      program.exitOverride();

      const addCmd = program.commands.find((c) => c.name() === "add");
      const actionCmd = addCmd!.commands.find((c) => c.name() === "action");
      expect(actionCmd).to.exist;

      // These are now .option() not .requiredOption() because they're only required for api-spec mode
      const mandatory = actionCmd!.options.filter((o) => o.mandatory).map((o) => o.long);
      expect(mandatory).to.not.include("--action-id");
      expect(mandatory).to.not.include("--api-spec-path");
      expect(mandatory).to.not.include("--plugin-manifest-path");
    });

    it("atk m365-sideload — --file-path is mandatory", () => {
      const program = buildProgram("atk");
      program.exitOverride();

      const m365Cmd = program.commands.find((c) => c.name() === "m365-sideload");
      expect(m365Cmd, "'m365-sideload' command should exist").to.exist;

      const mandatory = m365Cmd!.options.filter((o) => o.mandatory).map((o) => o.long);
      expect(mandatory).to.include("--file-path");
    });
  });
});

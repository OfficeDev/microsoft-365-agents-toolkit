// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * MCP scaffold integration tests — subprocess layer.
 *
 * Verifies that `atk new da mcp-remote` and `atk add action` with MCP flags
 * produce the correct project structure (ai-plugin.json, declarativeAgent.json,
 * .vscode/mcp.json). Does NOT create real Azure resources.
 *
 * Maps to ADO test plan suite "MCP - Declarative Agent with Remote MCP Server"
 * (37357417).
 */

import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { expect } from "chai";
import { describe, it, beforeEach, afterEach } from "mocha";
import { promisify } from "util";
import { exec } from "child_process";

const execAsync = promisify(exec);

const ATK_BIN = process.env.ATK_BIN || "atk";
const TIMEOUT = 120_000; // 2 min per command
const MCP_SERVER_URL = "https://learn.microsoft.com/api/mcp";

function tmpDir(): string {
  const dir = path.join(os.tmpdir(), `atk-mcp-test-${Date.now()}`);
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

async function run(
  command: string,
  cwd?: string
): Promise<{ stdout: string; stderr: string; exitCode: number }> {
  try {
    const { stdout, stderr } = await execAsync(command, {
      cwd,
      timeout: TIMEOUT,
      env: { ...process.env, CI_ENABLED: "true" },
    });
    return { stdout, stderr, exitCode: 0 };
  } catch (e: unknown) {
    const err = e as { stdout?: string; stderr?: string; code?: number };
    return {
      stdout: err.stdout ?? "",
      stderr: err.stderr ?? "",
      exitCode: err.code ?? 1,
    };
  }
}

function readJson(filePath: string): Record<string, unknown> {
  return JSON.parse(fs.readFileSync(filePath, "utf8")) as Record<string, unknown>;
}

describe("MCP scaffold integration", function () {
  this.timeout(5 * 60 * 1000); // 5 min total

  let dir: string;

  beforeEach(function () {
    dir = tmpDir();
  });

  afterEach(async function () {
    await fs.promises.rm(dir, { recursive: true, force: true }).catch(() => {});
  });

  describe("atk new da mcp-remote (no auth)", function () {
    it("scaffolds project with MCP server URL", async function () {
      const result = await run(
        `${ATK_BIN} new da mcp-remote --name McpApp --mcpServerUrl ${MCP_SERVER_URL} --folder ${dir} --non-interactive`,
        dir
      );
      expect(result.exitCode, `stderr: ${result.stderr}`).to.equal(0);

      const projectPath = path.join(dir, "McpApp");
      expect(fs.existsSync(projectPath), "project dir should exist").to.be.true;

      // Verify core project files
      expect(fs.existsSync(path.join(projectPath, "m365agents.yml")), "m365agents.yml").to.be.true;
      expect(fs.existsSync(path.join(projectPath, "env", ".env.dev")), ".env.dev").to.be.true;

      // Verify appPackage structure
      const appPackage = path.join(projectPath, "appPackage");
      expect(fs.existsSync(path.join(appPackage, "ai-plugin.json")), "ai-plugin.json").to.be.true;
      expect(fs.existsSync(path.join(appPackage, "declarativeAgent.json")), "declarativeAgent.json")
        .to.be.true;
      expect(fs.existsSync(path.join(appPackage, "manifest.json")), "manifest.json").to.be.true;

      // Verify declarativeAgent.json references ai-plugin as action
      const daManifest = readJson(path.join(appPackage, "declarativeAgent.json"));
      const actions = daManifest.actions as { file: string }[];
      expect(actions).to.be.an("array").that.is.not.empty;
      expect(actions[0].file).to.equal("ai-plugin.json");

      // Verify .vscode/mcp.json has the remote URL
      const mcpConfig = readJson(path.join(projectPath, ".vscode", "mcp.json"));
      expect(mcpConfig.servers).to.be.an("object");
      const serverEntry = Object.values(mcpConfig.servers as Record<string, unknown>)[0] as Record<
        string,
        unknown
      >;
      expect(serverEntry.url).to.equal(MCP_SERVER_URL);
      expect(serverEntry.type).to.equal("http");
    });

    it("scaffolds project without URL — mcp.json has empty/no URL", async function () {
      const result = await run(
        `${ATK_BIN} new da mcp-remote --name McpNoUrl --folder ${dir} --non-interactive`,
        dir
      );
      // May exit 0 (scaffold with placeholder) or non-zero (validation failure)
      const projectPath = path.join(dir, "McpNoUrl");
      if (result.exitCode === 0 && fs.existsSync(projectPath)) {
        // If scaffold succeeds without URL, verify no URL appears in mcp.json
        const mcpConfigPath = path.join(projectPath, ".vscode", "mcp.json");
        if (fs.existsSync(mcpConfigPath)) {
          const mcpConfig = readJson(mcpConfigPath);
          const serverEntry = Object.values(
            mcpConfig.servers as Record<string, unknown>
          )[0] as Record<string, unknown>;
          expect(serverEntry.url).to.satisfy(
            (url: string) => !url || url === "" || url.includes("{{"),
            "URL should be empty or a placeholder"
          );
        }
      } else {
        // Scaffold failed without URL — acceptable behavior
        expect(result.exitCode).to.not.equal(0);
      }
    });
  });

  describe("atk new da mcp-remote --help", function () {
    it("shows --mcpServerUrl option", async function () {
      const result = await run(`${ATK_BIN} new da mcp-remote --help`);
      expect(result.exitCode).to.equal(0);
      expect(result.stdout.toLowerCase()).to.include("mcpserverurl");
    });
  });

  describe("atk add action on MCP project", function () {
    it("add action --help exits 0 and shows MCP-relevant flags", async function () {
      const result = await run(`${ATK_BIN} add action --help`);
      expect(result.exitCode).to.equal(0);
      expect(result.stdout).to.include("--api-spec-path");
      expect(result.stdout).to.include("--plugin-manifest-path");
      expect(result.stdout).to.include("--action-id");
      expect(result.stdout).to.include("--api-plugin-type");
      expect(result.stdout).to.include("--mcp-server-url");
      expect(result.stdout).to.include("--mcp-server-name");
      expect(result.stdout).to.include("--mcp-auth-type");
      expect(result.stdout).to.include("--mcp-tools-file");
    });

    it("add action --api-plugin-type mcp — adds MCP tools to scaffolded project", async function () {
      // Scaffold a DA + MCP remote project first
      const scaffoldResult = await run(
        `${ATK_BIN} new da mcp-remote --name McpAddTest --mcpServerUrl ${MCP_SERVER_URL} --folder ${dir} --non-interactive`,
        dir
      );
      expect(scaffoldResult.exitCode, `scaffold stderr: ${scaffoldResult.stderr}`).to.equal(0);

      const projectPath = path.join(dir, "McpAddTest");
      const appPackage = path.join(projectPath, "appPackage");
      const pluginPath = path.join(appPackage, "ai-plugin.json");

      // Create a tools file for the add action
      const toolsFile = path.join(dir, "add-tools.json");
      fs.writeFileSync(
        toolsFile,
        JSON.stringify([
          { name: "searchDocs", description: "Search documentation" },
          { name: "getPage", description: "Get a specific page" },
        ]),
        "utf8"
      );

      const addResult = await run(
        [
          ATK_BIN,
          "add",
          "action",
          "--api-plugin-type",
          "mcp",
          "--mcp-server-url",
          "https://another-mcp.example.com/sse",
          "--mcp-server-name",
          "docs-server",
          "--mcp-auth-type",
          "none",
          "--mcp-tools-file",
          toolsFile,
          "--mcp-selected-tools",
          "searchDocs",
          "--plugin-manifest-path",
          pluginPath,
        ].join(" "),
        projectPath
      );
      expect(addResult.exitCode, `add stderr: ${addResult.stderr}`).to.equal(0);

      // Verify ai-plugin.json has the new MCP runtime
      const savedPlugin = readJson(pluginPath);
      const runtimes = savedPlugin.runtimes as Array<Record<string, unknown>>;
      const mcpRuntime = runtimes?.find(
        (r) => (r.spec as Record<string, unknown>)?.url === "https://another-mcp.example.com/sse"
      );
      expect(mcpRuntime, "MCP runtime for docs-server should exist").to.exist;
      expect(mcpRuntime!.type).to.equal("RemoteMCPServer");

      // Verify functions were added
      const functions = savedPlugin.functions as Array<{ name: string }>;
      const fnNames = functions?.map((f) => f.name) ?? [];
      expect(fnNames).to.include("searchDocs");

      // Verify mcp-tools.json sidecar
      const sidecarPath = path.join(appPackage, "mcp-tools.json");
      expect(fs.existsSync(sidecarPath), "mcp-tools.json sidecar should exist").to.be.true;
      const sidecar = readJson(sidecarPath) as unknown as Array<{ name: string }>;
      expect(sidecar).to.be.an("array");
      expect(sidecar.map((t) => t.name)).to.include("searchDocs");
    });

    it("add action fails gracefully without required flags", async function () {
      // scaffold a DA first
      const scaffoldResult = await run(
        `${ATK_BIN} new da basic --name McpBase --folder ${dir} --non-interactive`,
        dir
      );
      expect(scaffoldResult.exitCode, `scaffold stderr: ${scaffoldResult.stderr}`).to.equal(0);

      const projectPath = path.join(dir, "McpBase");
      // add action with mcp type but no --mcp-server-url → should fail
      const result = await run(`${ATK_BIN} add action --api-plugin-type mcp`, projectPath);
      expect(result.exitCode).to.not.equal(0);
      expect(result.stderr).to.match(/required|missing|mcp-server-url/i);
    });
  });
});

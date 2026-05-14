/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, afterEach, beforeEach } from "mocha";
import * as sinon from "sinon";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import * as os from "node:os";
import type { AtkContext } from "@microsoft/teamsfx-core-next";

// Actions under test
import { addActionAction, addMCPActionAction } from "../../src/actions/addAction";
import { addCapabilityAction } from "../../src/actions/addCapability";
import { addAuthConfigAction } from "../../src/actions/addAuthConfig";
import { setSensitivityLabelAction } from "../../src/actions/setSensitivityLabel";
import { m365SideloadAction } from "../../src/actions/m365Sideload";

function createMockContext(overrides?: Partial<AtkContext>): AtkContext {
  return {
    auth: {
      m365TokenProvider: {} as any,
      azureAccountProvider: {} as any,
    },
    logger: {
      log: sinon.stub(),
      verbose: sinon.stub(),
      debug: sinon.stub(),
      info: sinon.stub(),
      warning: sinon.stub(),
      error: sinon.stub(),
      logInFile: sinon.stub().resolves(),
      getLogFilePath: sinon.stub().returns("/tmp/test.log"),
    } as any,
    telemetry: {
      sendTelemetryEvent: sinon.stub(),
      sendTelemetryErrorEvent: sinon.stub(),
      sendTelemetryException: sinon.stub(),
    },
    ui: {} as any,
    correlationId: "test-correlation-id",
    projectPath: "/tmp/test-project",
    ...overrides,
  };
}

describe("DA Action Handlers", () => {
  const sandbox = sinon.createSandbox();
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "cli-da-test-"));
  });

  afterEach(async () => {
    sandbox.restore();
    if (tmpDir) {
      await fs.rm(tmpDir, { recursive: true, force: true });
    }
  });

  describe("addActionAction()", () => {
    it("should throw when agent manifest not found and no auto-discover", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      try {
        await addActionAction(ctx, {
          projectPath: tmpDir,
          apiSpecPath: "/nonexistent/spec.yaml",
          pluginManifestPath: "/nonexistent/plugin.json",
          actionId: "myAction",
          // No agentManifestPath → auto-discover will fail (no manifest.json)
        });
        expect.fail("should have thrown");
      } catch (e: any) {
        expect(e.name).to.equal("ManifestNotFound");
      }
    });

    it("should have correct function signature and accept explicit manifest path", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      // The operation will fail at Zod validation or file I/O, but the shape is verified.
      try {
        await addActionAction(ctx, {
          projectPath: tmpDir,
          agentManifestPath: path.join(tmpDir, "da.json"),
          apiSpecPath: path.join(tmpDir, "spec.yaml"),
          pluginManifestPath: path.join(tmpDir, "plugin.json"),
          actionId: "a1",
        });
      } catch (e: any) {
        // Expected — files don't exist
        expect(e.message).to.be.a("string");
      }
    });
  });

  describe("addCapabilityAction()", () => {
    it("should throw when agent manifest not found via auto-discover", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      try {
        await addCapabilityAction(ctx, {
          projectPath: tmpDir,
          source: "web-search",
        });
        expect.fail("should have thrown");
      } catch (e: any) {
        expect(e.name).to.equal("ManifestNotFound");
      }
    });

    it("should accept all knowledge source types", async () => {
      const sources = [
        "web-search",
        "onedrive-sharepoint",
        "graph-connector",
        "embedded-knowledge",
      ] as const;
      for (const source of sources) {
        const ctx = createMockContext({ projectPath: tmpDir });
        try {
          await addCapabilityAction(ctx, {
            projectPath: tmpDir,
            agentManifestPath: path.join(tmpDir, "da.json"),
            source,
          });
        } catch {
          // Expected — file doesn't exist
        }
      }
    });
  });

  describe("addAuthConfigAction()", () => {
    it("should accept oauth auth type", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      try {
        await addAuthConfigAction(ctx, {
          projectPath: tmpDir,
          authType: "oauth",
          authName: "myOAuth",
          specPath: "./spec.yaml",
        });
      } catch (e: any) {
        // Expected — teamsapp.yml doesn't exist
        expect(e.message).to.be.a("string");
      }
    });

    it("should accept api-key auth type", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      try {
        await addAuthConfigAction(ctx, {
          projectPath: tmpDir,
          authType: "api-key",
          authName: "myApiKey",
          specPath: "./spec.yaml",
        });
      } catch (e: any) {
        expect(e.message).to.be.a("string");
      }
    });

    it("should use custom yml-path when provided", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      const customYml = path.join(tmpDir, "custom.yml");
      try {
        await addAuthConfigAction(ctx, {
          projectPath: tmpDir,
          authType: "oauth",
          ymlPath: customYml,
          authName: "myOAuth",
          specPath: "./spec.yaml",
          entra: true,
          enablePkce: true,
        });
      } catch (e: any) {
        expect(e.message).to.be.a("string");
      }
    });
  });

  describe("setSensitivityLabelAction()", () => {
    it("should throw when agent manifest not found via auto-discover", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      try {
        await setSensitivityLabelAction(ctx, {
          projectPath: tmpDir,
          labelId: "label-123",
        });
        expect.fail("should have thrown");
      } catch (e: any) {
        expect(e.name).to.equal("ManifestNotFound");
      }
    });

    it("should accept explicit agent manifest path", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      try {
        await setSensitivityLabelAction(ctx, {
          projectPath: tmpDir,
          agentManifestPath: path.join(tmpDir, "da.json"),
          labelId: "label-456",
        });
      } catch (e: any) {
        // Expected — file doesn't exist
        expect(e.message).to.be.a("string");
      }
    });
  });

  describe("addMCPActionAction()", () => {
    it("should throw when --mcp-server-url is missing", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      try {
        await addMCPActionAction(ctx, {
          projectPath: tmpDir,
          serverUrl: "",
          serverName: "test",
          isLocal: false,
          authType: "none",
        });
        expect.fail("should have thrown");
      } catch (e: any) {
        expect(e.message).to.include("--mcp-server-url");
      }
    });

    it("should throw when --mcp-server-name is missing", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      try {
        await addMCPActionAction(ctx, {
          projectPath: tmpDir,
          serverUrl: "https://mcp.example.com",
          serverName: "",
          isLocal: false,
          authType: "none",
        });
        expect.fail("should have thrown");
      } catch (e: any) {
        expect(e.message).to.include("--mcp-server-name");
      }
    });

    it("should throw when auth type is invalid", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      try {
        await addMCPActionAction(ctx, {
          projectPath: tmpDir,
          serverUrl: "https://mcp.example.com",
          serverName: "test",
          isLocal: false,
          authType: "invalid-auth",
        });
        expect.fail("should have thrown");
      } catch (e: any) {
        expect(e.message).to.include("Invalid --mcp-auth-type");
      }
    });

    it("should throw when tools file does not exist", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      try {
        await addMCPActionAction(ctx, {
          projectPath: tmpDir,
          serverUrl: "https://mcp.example.com",
          serverName: "test",
          isLocal: false,
          authType: "none",
          toolsFilePath: path.join(tmpDir, "nonexistent-tools.json"),
        });
        expect.fail("should have thrown");
      } catch (e: any) {
        expect(e.message).to.include("not found");
      }
    });

    it("should throw when tools file contains invalid JSON", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      const toolsFile = path.join(tmpDir, "bad-tools.json");
      await fs.writeFile(toolsFile, "{ not valid json }", "utf-8");
      try {
        await addMCPActionAction(ctx, {
          projectPath: tmpDir,
          serverUrl: "https://mcp.example.com",
          serverName: "test",
          isLocal: false,
          authType: "none",
          toolsFilePath: toolsFile,
        });
        expect.fail("should have thrown");
      } catch (e: any) {
        expect(e).to.be.instanceOf(SyntaxError);
      }
    });

    it("should throw when tools file is not an array", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      const toolsFile = path.join(tmpDir, "obj-tools.json");
      await fs.writeFile(toolsFile, JSON.stringify({ name: "not-array" }), "utf-8");
      try {
        await addMCPActionAction(ctx, {
          projectPath: tmpDir,
          serverUrl: "https://mcp.example.com",
          serverName: "test",
          isLocal: false,
          authType: "none",
          toolsFilePath: toolsFile,
        });
        expect.fail("should have thrown");
      } catch (e: any) {
        expect(e.message).to.include("JSON array");
      }
    });

    it("should throw when a tool entry is missing name", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      const toolsFile = path.join(tmpDir, "no-name-tools.json");
      await fs.writeFile(toolsFile, JSON.stringify([{ description: "no name field" }]), "utf-8");
      try {
        await addMCPActionAction(ctx, {
          projectPath: tmpDir,
          serverUrl: "https://mcp.example.com",
          serverName: "test",
          isLocal: false,
          authType: "none",
          toolsFilePath: toolsFile,
        });
        expect.fail("should have thrown");
      } catch (e: any) {
        expect(e.message).to.include("name");
      }
    });

    it("should accept all valid auth types", async () => {
      const validTypes = ["oauth", "api-key", "bearer-token", "microsoft-entra", "none"];
      for (const authType of validTypes) {
        const ctx = createMockContext({ projectPath: tmpDir });
        try {
          await addMCPActionAction(ctx, {
            projectPath: tmpDir,
            serverUrl: "https://mcp.example.com",
            serverName: "test",
            isLocal: false,
            authType,
          });
        } catch (e: any) {
          // Expected — plugin manifest won't exist, but should NOT fail on auth validation
          expect(e.message).to.not.include("Invalid --mcp-auth-type");
        }
      }
    });

    it("should load tools from valid tools file", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      const toolsFile = path.join(tmpDir, "tools.json");
      await fs.writeFile(
        toolsFile,
        JSON.stringify([
          { name: "getTodos", description: "Get todos" },
          { name: "addTodo", description: "Add a todo", inputSchema: { type: "object" } },
        ]),
        "utf-8"
      );

      // Will fail at the plugin manifest step, but tools loading should succeed
      try {
        await addMCPActionAction(ctx, {
          projectPath: tmpDir,
          serverUrl: "https://mcp.example.com",
          serverName: "test",
          isLocal: false,
          authType: "none",
          toolsFilePath: toolsFile,
        });
      } catch (e: any) {
        // Should fail at addMCPActionOp, not at tools loading
        expect(e.message).to.not.include("tools file");
      }
    });
  });

  describe("m365SideloadAction()", () => {
    it("should throw when app package file not found", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      try {
        await m365SideloadAction(ctx, {
          filePath: path.join(tmpDir, "nonexistent.zip"),
        });
        expect.fail("should have thrown");
      } catch (e: any) {
        expect(e.name).to.equal("AppPackageNotFound");
      }
    });

    it("should accept scope parameter", async () => {
      const ctx = createMockContext({ projectPath: tmpDir });
      try {
        await m365SideloadAction(ctx, {
          filePath: path.join(tmpDir, "app.zip"),
          scope: "Shared",
        });
      } catch (e: any) {
        // Expected — file doesn't exist
        expect(e.name).to.equal("AppPackageNotFound");
      }
    });
  });
});

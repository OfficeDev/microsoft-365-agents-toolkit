/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, afterEach } from "mocha";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import * as os from "node:os";
import { runOperation } from "../../src/core/operation";
import {
  addKnowledgeOp,
  addExistingPluginOp,
  setSensitivityLabelOp,
  setConversationStartersOp,
} from "../../src/declarativeAgent/operations";
import { extendToM365Op } from "../../src/teamsApp/operations";
import { createMockContext } from "../unit/testHelper";

/**
 * Integration tests: DA operations on real filesystem.
 *
 * Each test creates a minimal DA project in a temp directory,
 * runs the operation via runOperation(), and verifies filesystem side effects.
 */
describe("integration: DA operations", () => {
  let tmpDir: string;

  afterEach(async () => {
    if (tmpDir) {
      await fs.rm(tmpDir, { recursive: true, force: true });
    }
  });

  /**
   * Create a minimal DA project with manifest.json + declarativeAgent.json + icons.
   */
  async function createDAProject(baseDir: string): Promise<{
    appDir: string;
    agentManifestPath: string;
  }> {
    const appDir = path.join(baseDir, "appPackage");
    await fs.mkdir(appDir, { recursive: true });

    // Icons
    const pngHeader = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    await fs.writeFile(path.join(appDir, "color.png"), pngHeader);
    await fs.writeFile(path.join(appDir, "outline.png"), pngHeader);

    // Teams manifest
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
    await fs.writeFile(path.join(appDir, "manifest.json"), JSON.stringify(manifest, null, 2));

    // DA manifest (minimal)
    const daManifest = {
      $schema:
        "https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.6/schema.json",
      version: "v1.6",
      name: "Test DA",
      description: "A test declarative agent",
      instructions: "Be helpful",
    };
    const agentManifestPath = path.join(appDir, "declarativeAgent.json");
    await fs.writeFile(agentManifestPath, JSON.stringify(daManifest, null, 2));

    return { appDir, agentManifestPath };
  }

  // ─── addKnowledge: web-search ──────────────────────────────

  it("addKnowledgeOp (web-search) — adds WebSearch capability to manifest", async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "atk-integ-daops-"));
    const { agentManifestPath } = await createDAProject(tmpDir);
    const ctx = createMockContext({ projectPath: tmpDir });

    const result = await runOperation(addKnowledgeOp, ctx, {
      agentManifestPath,
      source: "web-search",
    });

    expect(result.isOk(), `Failed: ${result.isErr() ? result.error.message : ""}`).to.be.true;

    const updated = JSON.parse(await fs.readFile(agentManifestPath, "utf-8"));
    const caps = updated.capabilities as Array<Record<string, unknown>>;
    expect(caps).to.be.an("array").with.length.greaterThan(0);
    const webSearch = caps.find((c) => c.name === "WebSearch");
    expect(webSearch, "WebSearch capability should be present").to.exist;
  });

  // ─── addKnowledge: embedded ────────────────────────────────

  it("addKnowledgeOp (embedded) — copies files and updates manifest", async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "atk-integ-daops-"));
    const { appDir, agentManifestPath } = await createDAProject(tmpDir);
    const ctx = createMockContext({ projectPath: tmpDir });

    // Create a source file to embed
    const srcFile = path.join(tmpDir, "faq.txt");
    await fs.writeFile(srcFile, "Q: What is this?\nA: A test.");

    const result = await runOperation(addKnowledgeOp, ctx, {
      agentManifestPath,
      source: "embedded-knowledge",
      embeddedFilePaths: [srcFile],
    });

    expect(result.isOk(), `Failed: ${result.isErr() ? result.error.message : ""}`).to.be.true;

    // Verify knowledge directory was created and file was copied
    const knowledgeDir = path.join(appDir, "knowledge");
    const copiedFile = path.join(knowledgeDir, "faq.txt");
    const exists = await fs
      .access(copiedFile)
      .then(() => true)
      .catch(() => false);
    expect(exists, "Embedded file should be copied to knowledge/").to.be.true;

    // Verify manifest was updated
    const updated = JSON.parse(await fs.readFile(agentManifestPath, "utf-8"));
    const caps = updated.capabilities as Array<Record<string, unknown>>;
    const embedded = caps.find((c) => c.name === "EmbeddedKnowledge");
    expect(embedded, "EmbeddedKnowledge capability should be present").to.exist;
  });

  // ─── addExistingPlugin ─────────────────────────────────────

  it("addExistingPluginOp — registers action and copies files", async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "atk-integ-daops-"));
    const { appDir, agentManifestPath } = await createDAProject(tmpDir);
    const ctx = createMockContext({ projectPath: tmpDir });

    // Create a plugin manifest
    const pluginManifest = {
      $schema: "https://aka.ms/json-schemas/copilot/plugin/v2.2/schema.json",
      schema_version: "v2.2",
      name_for_human: "Test Plugin",
      namespace: "test",
      description_for_human: "A test plugin",
      functions: [{ name: "getItems", description: "Get items" }],
      runtimes: [
        {
          type: "OpenApi",
          auth: { type: "None" },
          spec: { url: "openapi.yaml" },
          run_for_functions: ["getItems"],
        },
      ],
    };
    const pluginPath = path.join(tmpDir, "plugin.json");
    await fs.writeFile(pluginPath, JSON.stringify(pluginManifest, null, 2));

    // Create an OpenAPI spec
    const specPath = path.join(tmpDir, "openapi.yaml");
    await fs.writeFile(
      specPath,
      'openapi: "3.0.0"\ninfo:\n  title: Test\n  version: "1.0"\npaths:\n  /items:\n    get:\n      operationId: getItems\n      summary: Get items\n      responses:\n        "200":\n          description: OK\n'
    );

    const result = await runOperation(addExistingPluginOp, ctx, {
      agentManifestPath,
      pluginManifestPath: pluginPath,
      apiSpecPath: specPath,
      actionId: "testAction",
    });

    expect(result.isOk(), `Failed: ${result.isErr() ? result.error.message : ""}`).to.be.true;

    // Verify DA manifest now has the action registered
    const updated = JSON.parse(await fs.readFile(agentManifestPath, "utf-8"));
    const actions = updated.actions as Array<{ id: string; file: string }>;
    expect(actions).to.be.an("array").with.length.greaterThan(0);
    const registered = actions.find((a) => a.id === "testAction");
    expect(registered, "Action 'testAction' should be registered").to.exist;
  });

  // ─── setSensitivityLabel ───────────────────────────────────

  it("setSensitivityLabelOp — sets sensitivity_label.id in manifest", async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "atk-integ-daops-"));
    const { agentManifestPath } = await createDAProject(tmpDir);
    const ctx = createMockContext({ projectPath: tmpDir });

    const result = await runOperation(setSensitivityLabelOp, ctx, {
      agentManifestPath,
      labelId: "label-abc-123",
    });

    expect(result.isOk(), `Failed: ${result.isErr() ? result.error.message : ""}`).to.be.true;

    const updated = JSON.parse(await fs.readFile(agentManifestPath, "utf-8"));
    expect(updated.sensitivity_label).to.deep.equal({ id: "label-abc-123" });
  });

  // ─── setConversationStarters ───────────────────────────────

  it("setConversationStartersOp — adds starters to manifest", async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "atk-integ-daops-"));
    const { agentManifestPath } = await createDAProject(tmpDir);
    const ctx = createMockContext({ projectPath: tmpDir });

    const result = await runOperation(setConversationStartersOp, ctx, {
      agentManifestPath,
      starters: [
        { text: "How can I help?", title: "Greeting" },
        { text: "Tell me about the project" },
      ],
    });

    expect(result.isOk(), `Failed: ${result.isErr() ? result.error.message : ""}`).to.be.true;

    const updated = JSON.parse(await fs.readFile(agentManifestPath, "utf-8"));
    expect(updated.conversation_starters).to.be.an("array").with.length(2);
    expect(updated.conversation_starters[0].text).to.equal("How can I help?");
    expect(updated.conversation_starters[0].title).to.equal("Greeting");
  });

  // ─── extendToM365Op validation ─────────────────────────────

  it("extendToM365Op — returns AppPackageNotFound for missing file", async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "atk-integ-daops-"));
    const ctx = createMockContext({ projectPath: tmpDir });

    const result = await runOperation(extendToM365Op, ctx, {
      appPackagePath: path.join(tmpDir, "nonexistent.zip"),
    });

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("AppPackageNotFound");
    }
  });
});

/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, afterEach } from "mocha";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import * as os from "node:os";
import AdmZip from "adm-zip";
import { buildAppPackage } from "../../src/teamsApp/packageBuilder";
import { createMockContext } from "../unit/testHelper";

/**
 * Integration test: full Declarative Agent packaging pipeline.
 *
 * Creates a realistic project on disk with:
 * - manifest.json referencing a DA
 * - DA manifest with a plugin action
 * - Plugin manifest with OpenAPI spec runtime
 * - Embedded knowledge file
 * - Icons
 *
 * Then verifies the ZIP contains all expected entries.
 */
describe("integration: DA packaging pipeline", () => {
  let tmpDir: string;

  afterEach(async () => {
    if (tmpDir) {
      await fs.rm(tmpDir, { recursive: true, force: true });
    }
  });

  async function setup(): Promise<string> {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "atk-integ-dapkg-"));
    return tmpDir;
  }

  it("should package a full DA project with plugin, spec, and knowledge files", async () => {
    const dir = await setup();
    const appDir = path.join(dir, "appPackage");
    await fs.mkdir(appDir, { recursive: true });

    // ── Create icons ──
    const pngHeader = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    await fs.writeFile(path.join(appDir, "color.png"), pngHeader);
    await fs.writeFile(path.join(appDir, "outline.png"), pngHeader);

    // ── Create manifest.json ──
    const manifest = {
      $schema:
        "https://developer.microsoft.com/en-us/json-schemas/teams/v1.19/MicrosoftTeams.schema.json",
      manifestVersion: "1.19",
      version: "1.0.0",
      id: "${{TEAMS_APP_ID}}",
      developer: {
        name: "Test Corp",
        websiteUrl: "https://example.com",
        privacyUrl: "https://example.com/privacy",
        termsOfUseUrl: "https://example.com/tos",
      },
      name: { short: "${{APP_NAME}}", full: "DA Test App" },
      description: { short: "A DA app", full: "A declarative agent test app" },
      icons: { color: "color.png", outline: "outline.png" },
      accentColor: "#4F6BED",
      copilotAgents: {
        declarativeAgents: [
          {
            id: "daAgent",
            file: "declarativeAgent.json",
          },
        ],
      },
    };
    await fs.writeFile(path.join(appDir, "manifest.json"), JSON.stringify(manifest, null, 2));

    // ── Create DA manifest ──
    const daManifest = {
      $schema:
        "https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.6/schema.json",
      version: "v1.6",
      name: "Test DA",
      description: "A test declarative agent",
      instructions: "Be helpful",
      actions: [
        {
          id: "plugin1",
          file: "plugin.json",
        },
      ],
      capabilities: [
        {
          name: "EmbeddedKnowledge",
          files: [{ file: "knowledge/faq.txt" }],
        },
      ],
    };
    await fs.writeFile(
      path.join(appDir, "declarativeAgent.json"),
      JSON.stringify(daManifest, null, 2)
    );

    // ── Create plugin manifest ──
    const pluginManifest = {
      $schema: "https://aka.ms/json-schemas/copilot/plugin/v2.2/schema.json",
      schema_version: "v2.2",
      name_for_human: "Test Plugin",
      namespace: "test_plugin",
      description_for_human: "A test plugin",
      functions: [
        {
          name: "getItems",
          description: "Get items",
        },
      ],
      runtimes: [
        {
          type: "OpenApi",
          auth: { type: "None" },
          spec: {
            url: "openapi.yaml",
          },
          run_for_functions: ["getItems"],
        },
      ],
    };
    await fs.writeFile(path.join(appDir, "plugin.json"), JSON.stringify(pluginManifest, null, 2));

    // ── Create OpenAPI spec ──
    const openApiSpec = `openapi: "3.0.0"
info:
  title: Test API
  version: "1.0"
paths:
  /items:
    get:
      operationId: getItems
      summary: Get items
      responses:
        "200":
          description: OK
`;
    await fs.writeFile(path.join(appDir, "openapi.yaml"), openApiSpec);

    // ── Create embedded knowledge file ──
    const knowledgeDir = path.join(appDir, "knowledge");
    await fs.mkdir(knowledgeDir, { recursive: true });
    await fs.writeFile(path.join(knowledgeDir, "faq.txt"), "Q: What is this?\nA: A test DA.");

    // ── Build the package ──
    const outputZip = path.join(dir, "build", "appPackage.zip");
    const outputFolder = path.join(dir, "build");
    const ctx = createMockContext({ projectPath: dir });

    const result = await buildAppPackage(ctx, {
      projectPath: dir,
      outputZipPath: outputZip,
      outputFolder,
      envs: {
        TEAMS_APP_ID: "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",
        APP_NAME: "MyDAAgent",
        TEAMSFX_ENV: "dev",
      },
    });

    // ── Verify ──
    expect(result.isOk(), `Build failed: ${result.isErr() ? (result.error as any).message : ""}`).to
      .be.true;
    if (result.isErr()) return;

    const { zipPath, jsonPath, fileCount } = result.value;

    // ZIP exists
    const zipStat = await fs.stat(zipPath);
    expect(zipStat.size).to.be.greaterThan(0);

    // Resolved JSON exists
    const jsonContent = await fs.readFile(jsonPath, "utf8");
    const parsedManifest = JSON.parse(jsonContent);
    expect(parsedManifest.id).to.equal("aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee");
    expect(parsedManifest.name.short).to.equal("MyDAAgent");

    // Verify ZIP contents
    const zip = new AdmZip(zipPath);
    const entries = zip.getEntries().map((e) => e.entryName);

    expect(entries).to.include("manifest.json");
    expect(entries).to.include("color.png");
    expect(entries).to.include("outline.png");
    expect(entries).to.include("declarativeAgent.json");
    expect(entries).to.include("plugin.json");
    expect(entries).to.include("openapi.yaml");
    expect(entries).to.include("knowledge/faq.txt");

    // Verify manifest.json in ZIP has resolved env vars
    const zipManifest = zip.getEntry("manifest.json");
    const zipManifestContent = JSON.parse(zipManifest!.getData().toString("utf8"));
    expect(zipManifestContent.id).to.equal("aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee");
    expect(zipManifestContent.name.short).to.equal("MyDAAgent");

    // Verify plugin namespace has underscore stripped
    const pluginEntry = zip.getEntry("plugin.json");
    const pluginContent = JSON.parse(pluginEntry!.getData().toString("utf8"));
    expect(pluginContent.namespace).to.equal("testplugin");

    // Verify knowledge file content
    const faqEntry = zip.getEntry("knowledge/faq.txt");
    expect(faqEntry).to.not.be.null;
    expect(faqEntry!.getData().toString("utf8")).to.contain("What is this?");

    expect(fileCount).to.be.greaterThanOrEqual(7);
  });
});

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
import { buildAppPackage } from "../../../src/teamsApp/packageBuilder";
import { createMockContext } from "../testHelper";

describe("teamsApp/packageBuilder", () => {
  let tmpDir: string;

  afterEach(async () => {
    if (tmpDir) {
      await fs.rm(tmpDir, { recursive: true, force: true });
    }
  });

  async function setup(): Promise<string> {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "atk-test-pkg-"));
    return tmpDir;
  }

  function makeManifest(overrides?: Record<string, unknown>): string {
    return JSON.stringify({
      $schema:
        "https://developer.microsoft.com/en-us/json-schemas/teams/v1.19/MicrosoftTeams.schema.json",
      manifestVersion: "1.19",
      version: "1.0.0",
      id: "00000000-0000-0000-0000-000000000000",
      developer: {
        name: "Test",
        websiteUrl: "https://example.com",
        privacyUrl: "https://example.com/privacy",
        termsOfUseUrl: "https://example.com/tos",
      },
      name: { short: "TestApp", full: "Test Application" },
      description: { short: "A test app", full: "A test application" },
      icons: { color: "color.png", outline: "outline.png" },
      accentColor: "#FFFFFF",
      ...overrides,
    });
  }

  async function createProjectDir(
    dir: string,
    manifestOverrides?: Record<string, unknown>
  ): Promise<string> {
    const appDir = path.join(dir, "appPackage");
    await fs.mkdir(appDir, { recursive: true });
    await fs.writeFile(path.join(appDir, "manifest.json"), makeManifest(manifestOverrides));

    // Create icon files (1x1 PNG placeholder)
    const pngHeader = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    await fs.writeFile(path.join(appDir, "color.png"), pngHeader);
    await fs.writeFile(path.join(appDir, "outline.png"), pngHeader);

    return appDir;
  }

  // ─── Basic packaging ─────────────────────────────────────

  it("should produce a valid ZIP with manifest and icons", async () => {
    const dir = await setup();
    await createProjectDir(dir);
    const outputZip = path.join(dir, "build", "appPackage.zip");
    const ctx = createMockContext({ projectPath: dir });

    const result = await buildAppPackage(ctx, {
      projectPath: dir,
      outputZipPath: outputZip,
    });

    expect(result.isOk()).to.be.true;
    if (result.isOk()) {
      expect(result.value.fileCount).to.be.greaterThanOrEqual(3);
      expect(result.value.zipPath).to.equal(outputZip);

      const zip = new AdmZip(outputZip);
      const entries = zip.getEntries().map((e) => e.entryName);
      expect(entries).to.include("manifest.json");
      expect(entries).to.include("color.png");
      expect(entries).to.include("outline.png");
    }
  });

  it("should write resolved manifest JSON alongside ZIP", async () => {
    const dir = await setup();
    await createProjectDir(dir);
    const outputZip = path.join(dir, "build", "appPackage.zip");
    const ctx = createMockContext({ projectPath: dir });

    const result = await buildAppPackage(ctx, {
      projectPath: dir,
      outputZipPath: outputZip,
    });

    expect(result.isOk()).to.be.true;
    if (result.isOk()) {
      // JSON file should exist
      const exists = await fs.access(result.value.jsonPath).then(
        () => true,
        () => false
      );
      expect(exists).to.be.true;

      const json = await fs.readFile(result.value.jsonPath, "utf8");
      const parsed = JSON.parse(json);
      expect(parsed.name.short).to.equal("TestApp");
    }
  });

  // ─── Env var resolution ──────────────────────────────────

  it("should resolve env vars in manifest", async () => {
    const dir = await setup();
    const appDir = path.join(dir, "appPackage");
    await fs.mkdir(appDir, { recursive: true });

    const manifest = JSON.stringify({
      $schema:
        "https://developer.microsoft.com/en-us/json-schemas/teams/v1.19/MicrosoftTeams.schema.json",
      manifestVersion: "1.19",
      version: "1.0.0",
      id: "${{APP_ID}}",
      developer: {
        name: "Test",
        websiteUrl: "https://example.com",
        privacyUrl: "https://example.com/privacy",
        termsOfUseUrl: "https://example.com/tos",
      },
      name: { short: "${{APP_NAME}}", full: "Test Application" },
      description: { short: "A test app", full: "A test application" },
      icons: { color: "color.png", outline: "outline.png" },
      accentColor: "#FFFFFF",
    });
    await fs.writeFile(path.join(appDir, "manifest.json"), manifest);

    const pngHeader = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    await fs.writeFile(path.join(appDir, "color.png"), pngHeader);
    await fs.writeFile(path.join(appDir, "outline.png"), pngHeader);

    const outputZip = path.join(dir, "build", "appPackage.zip");
    const ctx = createMockContext({ projectPath: dir });

    const result = await buildAppPackage(ctx, {
      projectPath: dir,
      outputZipPath: outputZip,
      envs: {
        APP_ID: "11111111-1111-1111-1111-111111111111",
        APP_NAME: "ResolvedApp",
      },
    });

    expect(result.isOk()).to.be.true;
    if (result.isOk()) {
      const zip = new AdmZip(outputZip);
      const manifestEntry = zip.getEntry("manifest.json");
      const content = manifestEntry!.getData().toString("utf8");
      const parsed = JSON.parse(content);
      expect(parsed.id).to.equal("11111111-1111-1111-1111-111111111111");
      expect(parsed.name.short).to.equal("ResolvedApp");
    }
  });

  // ─── Icon validation ─────────────────────────────────────

  it("should error when icon file is missing", async () => {
    const dir = await setup();
    const appDir = path.join(dir, "appPackage");
    await fs.mkdir(appDir, { recursive: true });
    await fs.writeFile(path.join(appDir, "manifest.json"), makeManifest());
    // Only create color.png, not outline.png
    const pngHeader = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    await fs.writeFile(path.join(appDir, "color.png"), pngHeader);

    const outputZip = path.join(dir, "build", "appPackage.zip");
    const ctx = createMockContext({ projectPath: dir });

    const result = await buildAppPackage(ctx, {
      projectPath: dir,
      outputZipPath: outputZip,
    });

    expect(result.isErr()).to.be.true;
  });

  // ─── Path traversal protection ───────────────────────────

  it("should reject icons outside app directory", async () => {
    const dir = await setup();
    const appDir = path.join(dir, "appPackage");
    await fs.mkdir(appDir, { recursive: true });

    // Manifest references icon outside app dir via ..
    const manifest = makeManifest({
      icons: { color: "../../../etc/passwd", outline: "outline.png" },
    });
    await fs.writeFile(path.join(appDir, "manifest.json"), manifest);
    const pngHeader = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    await fs.writeFile(path.join(appDir, "outline.png"), pngHeader);

    const outputZip = path.join(dir, "build", "appPackage.zip");
    const ctx = createMockContext({ projectPath: dir });

    const result = await buildAppPackage(ctx, {
      projectPath: dir,
      outputZipPath: outputZip,
    });

    expect(result.isErr()).to.be.true;
  });

  // ─── Explicit manifest path ──────────────────────────────

  it("should accept an explicit manifest path", async () => {
    const dir = await setup();
    const customDir = path.join(dir, "custom");
    await fs.mkdir(customDir, { recursive: true });

    await fs.writeFile(path.join(customDir, "manifest.json"), makeManifest());
    const pngHeader = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    await fs.writeFile(path.join(customDir, "color.png"), pngHeader);
    await fs.writeFile(path.join(customDir, "outline.png"), pngHeader);

    const outputZip = path.join(dir, "build", "appPackage.zip");
    const ctx = createMockContext({ projectPath: dir });

    const result = await buildAppPackage(ctx, {
      projectPath: dir,
      manifestPath: path.join(customDir, "manifest.json"),
      outputZipPath: outputZip,
    });

    expect(result.isOk()).to.be.true;
  });

  // ─── .generated folder preference ────────────────────────

  it("should prefer .generated folder when present", async () => {
    const dir = await setup();
    const appDir = path.join(dir, "appPackage");
    const genDir = path.join(appDir, ".generated");
    await fs.mkdir(genDir, { recursive: true });

    // Put different app names in each location
    await fs.writeFile(
      path.join(appDir, "manifest.json"),
      makeManifest({ name: { short: "Original", full: "Original App" } })
    );
    await fs.writeFile(
      path.join(genDir, "manifest.json"),
      makeManifest({ name: { short: "Generated", full: "Generated App" } })
    );

    const pngHeader = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    // Icons should be found relative to appDir (the directory containing .generated)
    await fs.writeFile(path.join(appDir, "color.png"), pngHeader);
    await fs.writeFile(path.join(appDir, "outline.png"), pngHeader);

    const outputZip = path.join(dir, "build", "appPackage.zip");
    const ctx = createMockContext({ projectPath: dir });

    const result = await buildAppPackage(ctx, {
      projectPath: dir,
      outputZipPath: outputZip,
    });

    expect(result.isOk()).to.be.true;
    if (result.isOk()) {
      const zip = new AdmZip(outputZip);
      const entry = zip.getEntry("manifest.json");
      const parsed = JSON.parse(entry!.getData().toString("utf8"));
      expect(parsed.name.short).to.equal("Generated");
    }
  });

  // ─── No manifest found ──────────────────────────────────

  it("should error when no manifest.json found", async () => {
    const dir = await setup();
    // Empty project directory
    const outputZip = path.join(dir, "build", "appPackage.zip");
    const ctx = createMockContext({ projectPath: dir });

    const result = await buildAppPackage(ctx, {
      projectPath: dir,
      outputZipPath: outputZip,
    });

    expect(result.isErr()).to.be.true;
  });
});

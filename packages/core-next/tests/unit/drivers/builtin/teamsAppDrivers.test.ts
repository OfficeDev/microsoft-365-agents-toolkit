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
import { DriverRegistry } from "../../../../src/drivers/registry";
import { zipAppPackageDriver } from "../../../../src/drivers/builtin/teamsApp/zipAppPackage";
import { validateManifestDriver } from "../../../../src/drivers/builtin/teamsApp/validateManifest";
import { validateAppPackageDriver } from "../../../../src/drivers/builtin/teamsApp/validateAppPackage";
import { builtinDrivers } from "../../../../src/drivers/builtin";
import { createMockContext } from "../../testHelper";

describe("teamsApp drivers", () => {
  let tmpDir: string;

  afterEach(async () => {
    if (tmpDir) {
      await fs.rm(tmpDir, { recursive: true, force: true });
    }
  });

  async function setup(): Promise<string> {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "atk-test-drivers-"));
    return tmpDir;
  }

  function makeManifest(): string {
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
    });
  }

  async function createProjectDir(dir: string): Promise<string> {
    const appDir = path.join(dir, "appPackage");
    await fs.mkdir(appDir, { recursive: true });
    await fs.writeFile(path.join(appDir, "manifest.json"), makeManifest());
    const pngHeader = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    await fs.writeFile(path.join(appDir, "color.png"), pngHeader);
    await fs.writeFile(path.join(appDir, "outline.png"), pngHeader);
    return appDir;
  }

  // ─── Registration ────────────────────────────────────────

  describe("registration", () => {
    it("should include all 13 drivers in builtinDrivers", () => {
      expect(builtinDrivers).to.have.lengthOf(22);
    });

    it("should include teamsApp driver IDs", () => {
      const ids = builtinDrivers.map((d) => d.id);
      expect(ids).to.include("teamsApp/zipAppPackage");
      expect(ids).to.include("teamsApp/validateManifest");
      expect(ids).to.include("teamsApp/validateAppPackage");
    });

    it("should register all 13 drivers into a fresh registry", () => {
      const registry = new DriverRegistry();
      for (const driver of builtinDrivers) {
        registry.register(driver);
      }
      expect(registry.size).to.equal(22);
      expect(registry.has("teamsApp/zipAppPackage")).to.be.true;
      expect(registry.has("teamsApp/validateManifest")).to.be.true;
      expect(registry.has("teamsApp/validateAppPackage")).to.be.true;
      expect(registry.has("teamsApp/create")).to.be.true;
      expect(registry.has("teamsApp/configure")).to.be.true;
      expect(registry.has("teamsApp/publishAppPackage")).to.be.true;
    });
  });

  // ─── zipAppPackage driver ────────────────────────────────

  describe("teamsApp/zipAppPackage", () => {
    it("should have correct driver metadata", () => {
      expect(zipAppPackageDriver.id).to.equal("teamsApp/zipAppPackage");
      expect(zipAppPackageDriver.name).to.equal("Zip App Package");
    });

    it("should produce a ZIP and output TEAMS_APP_PACKAGE_PATH", async () => {
      const dir = await setup();
      await createProjectDir(dir);
      const outputZip = path.join(dir, "build", "appPackage.zip");
      const ctx = createMockContext({ projectPath: dir });

      const result = await zipAppPackageDriver.executeFn(ctx, {
        projectPath: dir,
        outputZipPath: outputZip,
      });

      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.outputs).to.have.property("TEAMS_APP_PACKAGE_PATH");
        const zipPath = result.value.outputs.TEAMS_APP_PACKAGE_PATH;
        const stat = await fs.stat(zipPath);
        expect(stat.size).to.be.greaterThan(0);
      }
    });
  });

  // ─── validateManifest driver ─────────────────────────────

  describe("teamsApp/validateManifest", () => {
    it("should have correct driver metadata", () => {
      expect(validateManifestDriver.id).to.equal("teamsApp/validateManifest");
      expect(validateManifestDriver.name).to.equal("Validate Manifest");
    });

    it("should validate a correct manifest as true", async () => {
      const dir = await setup();
      const manifestPath = path.join(dir, "manifest.json");
      await fs.writeFile(manifestPath, makeManifest());
      const ctx = createMockContext({ projectPath: dir });

      const result = await validateManifestDriver.executeFn(ctx, {
        manifestPath,
      });

      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.outputs).to.have.property("TEAMS_APP_MANIFEST_VALID", "true");
      }
    });
  });

  // ─── validateAppPackage driver ───────────────────────────

  describe("teamsApp/validateAppPackage", () => {
    it("should have correct driver metadata", () => {
      expect(validateAppPackageDriver.id).to.equal("teamsApp/validateAppPackage");
      expect(validateAppPackageDriver.name).to.equal("Validate App Package");
    });

    it("should validate a correct package as true", async () => {
      const dir = await setup();
      await createProjectDir(dir);

      // First build a package
      const outputZip = path.join(dir, "build", "appPackage.zip");
      const ctx = createMockContext({ projectPath: dir });
      const buildResult = await zipAppPackageDriver.executeFn(ctx, {
        projectPath: dir,
        outputZipPath: outputZip,
      });
      expect(buildResult.isOk()).to.be.true;

      // Then validate it
      const result = await validateAppPackageDriver.executeFn(ctx, {
        appPackagePath: outputZip,
      });

      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.outputs).to.have.property("TEAMS_APP_PACKAGE_VALID", "true");
      }
    });

    it("should return error for missing package", async () => {
      const ctx = createMockContext();
      const result = await validateAppPackageDriver.executeFn(ctx, {
        appPackagePath: "/nonexistent/package.zip",
      });
      expect(result.isErr()).to.be.true;
    });

    it("should report false for package without manifest.json", async () => {
      const dir = await setup();
      const zipPath = path.join(dir, "empty.zip");
      const zip = new AdmZip();
      zip.addFile("readme.txt", Buffer.from("no manifest here"));
      zip.writeZip(zipPath);

      const ctx = createMockContext();
      const result = await validateAppPackageDriver.executeFn(ctx, {
        appPackagePath: zipPath,
      });

      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.outputs).to.have.property("TEAMS_APP_PACKAGE_VALID", "false");
      }
    });
  });
});

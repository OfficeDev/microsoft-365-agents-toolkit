/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, afterEach } from "mocha";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import * as os from "node:os";
import { getManifestPath, readTeamsManifest } from "../../../src/manifest/readManifest";

describe("manifest/readManifest", () => {
  let tmpDir: string;

  afterEach(async () => {
    if (tmpDir) {
      await fs.rm(tmpDir, { recursive: true, force: true });
    }
  });

  async function setup(): Promise<string> {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "atk-test-readmanifest-"));
    return tmpDir;
  }

  const minimalManifest = JSON.stringify({
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
    description: { short: "A test app", full: "A test application for testing" },
    icons: { color: "color.png", outline: "outline.png" },
    accentColor: "#FFFFFF",
  });

  // ─── getManifestPath ─────────────────────────────────────

  describe("getManifestPath", () => {
    it("should find manifest in appPackage/.generated", async () => {
      const dir = await setup();
      const genDir = path.join(dir, "appPackage", ".generated");
      await fs.mkdir(genDir, { recursive: true });
      await fs.writeFile(path.join(genDir, "manifest.json"), minimalManifest);

      const result = getManifestPath(dir);
      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value).to.equal(path.join(genDir, "manifest.json"));
      }
    });

    it("should find manifest in appPackage", async () => {
      const dir = await setup();
      const appDir = path.join(dir, "appPackage");
      await fs.mkdir(appDir, { recursive: true });
      await fs.writeFile(path.join(appDir, "manifest.json"), minimalManifest);

      const result = getManifestPath(dir);
      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value).to.equal(path.join(appDir, "manifest.json"));
      }
    });

    it("should find manifest at project root", async () => {
      const dir = await setup();
      await fs.writeFile(path.join(dir, "manifest.json"), minimalManifest);

      const result = getManifestPath(dir);
      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value).to.equal(path.join(dir, "manifest.json"));
      }
    });

    it("should prefer .generated over appPackage", async () => {
      const dir = await setup();
      const genDir = path.join(dir, "appPackage", ".generated");
      const appDir = path.join(dir, "appPackage");
      await fs.mkdir(genDir, { recursive: true });
      await fs.writeFile(path.join(genDir, "manifest.json"), minimalManifest);
      await fs.writeFile(path.join(appDir, "manifest.json"), minimalManifest);

      const result = getManifestPath(dir);
      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value).to.equal(path.join(genDir, "manifest.json"));
      }
    });

    it("should return error when no manifest found", async () => {
      const dir = await setup();
      const result = getManifestPath(dir);
      expect(result.isErr()).to.be.true;
    });
  });

  // ─── readTeamsManifest ───────────────────────────────────

  describe("readTeamsManifest", () => {
    it("should read and parse a valid manifest", async () => {
      const dir = await setup();
      const manifestPath = path.join(dir, "manifest.json");
      await fs.writeFile(manifestPath, minimalManifest);

      const result = await readTeamsManifest(manifestPath);
      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.wrapper.name.short).to.equal("TestApp");
        expect(result.value.raw).to.be.a("string");
      }
    });

    it("should strip BOM from manifest", async () => {
      const dir = await setup();
      const manifestPath = path.join(dir, "manifest.json");
      await fs.writeFile(manifestPath, "\uFEFF" + minimalManifest);

      const result = await readTeamsManifest(manifestPath);
      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.raw.charCodeAt(0)).to.not.equal(0xfeff);
      }
    });

    it("should return error for missing file", async () => {
      const result = await readTeamsManifest("/nonexistent/manifest.json");
      expect(result.isErr()).to.be.true;
    });
  });
});

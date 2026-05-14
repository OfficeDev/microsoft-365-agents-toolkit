/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, afterEach } from "mocha";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import * as os from "node:os";
import { getAgentManifestPath } from "../../../src/declarativeAgent/manifest/discovery";

describe("declarativeAgent/manifest/discovery", () => {
  let tmpDir: string;

  afterEach(async () => {
    if (tmpDir) {
      await fs.rm(tmpDir, { recursive: true, force: true });
    }
  });

  async function setup(): Promise<string> {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "atk-test-da-discovery-"));
    return tmpDir;
  }

  function makeTeamsManifest(daFile?: string): string {
    const base: Record<string, unknown> = {
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
    };
    if (daFile) {
      base.copilotAgents = {
        declarativeAgents: [{ id: "da1", file: daFile }],
      };
    }
    return JSON.stringify(base);
  }

  describe("getAgentManifestPath", () => {
    it("should discover DA manifest from appPackage/manifest.json", async () => {
      const dir = await setup();
      const appPackageDir = path.join(dir, "appPackage");
      await fs.mkdir(appPackageDir, { recursive: true });
      await fs.writeFile(
        path.join(appPackageDir, "manifest.json"),
        makeTeamsManifest("declarativeAgent.json")
      );

      const result = await getAgentManifestPath(dir);
      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.equal(path.join(appPackageDir, "declarativeAgent.json"));
    });

    it("should return error when no manifest.json found", async () => {
      const dir = await setup();

      const result = await getAgentManifestPath(dir);
      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("ManifestNotFound");
    });

    it("should return error when manifest has no declarative agents", async () => {
      const dir = await setup();
      const appPackageDir = path.join(dir, "appPackage");
      await fs.mkdir(appPackageDir, { recursive: true });
      await fs.writeFile(
        path.join(appPackageDir, "manifest.json"),
        makeTeamsManifest() // no DA
      );

      const result = await getAgentManifestPath(dir);
      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("NoDeclarativeAgent");
    });
  });
});

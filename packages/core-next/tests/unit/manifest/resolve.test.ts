/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, afterEach } from "mocha";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import * as os from "node:os";
import {
  resolveEnvPlaceholders,
  getEnvVariables,
  expandFunctionExpressions,
  resolveManifest,
} from "../../../src/manifest/resolve";
import { ManifestType } from "../../../src/manifest/types";
import { createMockContext } from "../testHelper";

describe("manifest/resolve", () => {
  let tmpDir: string;

  afterEach(async () => {
    if (tmpDir) {
      await fs.rm(tmpDir, { recursive: true, force: true });
    }
  });

  async function setup(): Promise<string> {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "atk-test-resolve-"));
    return tmpDir;
  }

  // ─── resolveEnvPlaceholders ──────────────────────────────

  describe("resolveEnvPlaceholders", () => {
    it("should replace known env vars", () => {
      const result = resolveEnvPlaceholders("Hello ${{NAME}}, your ID is ${{ID}}", {
        NAME: "Alice",
        ID: "42",
      });
      expect(result.content).to.equal("Hello Alice, your ID is 42");
      expect(result.unresolved).to.have.lengthOf(0);
    });

    it("should track unresolved placeholders", () => {
      const result = resolveEnvPlaceholders("${{KNOWN}} and ${{UNKNOWN}}", { KNOWN: "yes" });
      expect(result.content).to.equal("yes and ${{UNKNOWN}}");
      expect(result.unresolved).to.have.lengthOf(1);
      expect(result.unresolved[0].name).to.equal("UNKNOWN");
    });

    it("should return content unchanged when no placeholders", () => {
      const result = resolveEnvPlaceholders("plain text", {});
      expect(result.content).to.equal("plain text");
      expect(result.unresolved).to.have.lengthOf(0);
    });

    it("should handle empty content", () => {
      const result = resolveEnvPlaceholders("", { FOO: "bar" });
      expect(result.content).to.equal("");
    });

    it("should fall back to process.env when envs map omitted", () => {
      const key = "ATK_TEST_RESOLVE_FALLBACK_" + Date.now();
      process.env[key] = "from-process";
      try {
        const result = resolveEnvPlaceholders(`$\{{${key}}}`);
        expect(result.content).to.equal("from-process");
      } finally {
        delete process.env[key];
      }
    });
  });

  // ─── getEnvVariables ─────────────────────────────────────

  describe("getEnvVariables", () => {
    it("should extract unique variable names", () => {
      const vars = getEnvVariables("${{A}}, ${{B}}, ${{A}}");
      expect(vars).to.have.members(["A", "B"]);
      expect(vars).to.have.lengthOf(2);
    });

    it("should return empty array for no vars", () => {
      expect(getEnvVariables("no variables here")).to.have.lengthOf(0);
    });
  });

  // ─── expandFunctionExpressions ───────────────────────────

  describe("expandFunctionExpressions", () => {
    it("should expand $[file('path')] with static path", async () => {
      const dir = await setup();
      const dataFile = path.join(dir, "data.txt");
      await fs.writeFile(dataFile, "file-content");

      const manifestPath = path.join(dir, "manifest.json");
      const ctx = createMockContext({ projectPath: dir });

      const result = await expandFunctionExpressions(
        "value: \"$[file('data.txt')]\"",
        ctx,
        undefined,
        true,
        ManifestType.TeamsManifest,
        manifestPath
      );
      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value).to.equal('value: "file-content"');
      }
    });

    it("should expand $[file(${{VAR}})] with env var path", async () => {
      const dir = await setup();
      const dataFile = path.join(dir, "envfile.txt");
      await fs.writeFile(dataFile, "env-content");

      const manifestPath = path.join(dir, "manifest.json");
      const ctx = createMockContext({ projectPath: dir });

      const result = await expandFunctionExpressions(
        "$[file(${{FILE_PATH}})]",
        ctx,
        { FILE_PATH: "envfile.txt" },
        false,
        ManifestType.TeamsManifest,
        manifestPath
      );
      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value).to.equal("env-content");
      }
    });

    it("should error on unsupported file extension", async () => {
      const dir = await setup();
      const dataFile = path.join(dir, "data.json");
      await fs.writeFile(dataFile, "{}");

      const manifestPath = path.join(dir, "manifest.json");
      const ctx = createMockContext({ projectPath: dir });

      const result = await expandFunctionExpressions(
        "$[file('data.json')]",
        ctx,
        undefined,
        false,
        ManifestType.TeamsManifest,
        manifestPath
      );
      expect(result.isErr()).to.be.true;
    });

    it("should error on unsupported function name", async () => {
      const dir = await setup();
      const manifestPath = path.join(dir, "manifest.json");
      const ctx = createMockContext({ projectPath: dir });

      const result = await expandFunctionExpressions(
        "$[eval('something')]",
        ctx,
        undefined,
        false,
        ManifestType.TeamsManifest,
        manifestPath
      );
      expect(result.isErr()).to.be.true;
    });

    it("should return content unchanged when no functions", async () => {
      const ctx = createMockContext();
      const result = await expandFunctionExpressions(
        "no functions here",
        ctx,
        undefined,
        false,
        ManifestType.TeamsManifest,
        "/tmp/fake.json"
      );
      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value).to.equal("no functions here");
      }
    });

    it("should strip BOM from included files", async () => {
      const dir = await setup();
      const dataFile = path.join(dir, "bom.txt");
      await fs.writeFile(dataFile, "\uFEFFbom-content");

      const manifestPath = path.join(dir, "manifest.json");
      const ctx = createMockContext({ projectPath: dir });

      const result = await expandFunctionExpressions(
        "$[file('bom.txt')]",
        ctx,
        undefined,
        false,
        ManifestType.TeamsManifest,
        manifestPath
      );
      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value).to.equal("bom-content");
      }
    });
  });

  // ─── resolveManifest ─────────────────────────────────────

  describe("resolveManifest", () => {
    it("should resolve env vars in strict mode", async () => {
      const ctx = createMockContext();
      const result = await resolveManifest('{"name": "${{APP_NAME}}"}', ctx, {
        envs: { APP_NAME: "MyApp" },
        strict: true,
      });
      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value).to.equal('{"name": "MyApp"}');
      }
    });

    it("should error in strict mode on unresolved vars", async () => {
      const ctx = createMockContext();
      const result = await resolveManifest('{"name": "${{MISSING}}"}', ctx, {
        envs: {},
        strict: true,
      });
      expect(result.isErr()).to.be.true;
    });

    it("should allow unresolved vars in lenient mode", async () => {
      const ctx = createMockContext();
      const result = await resolveManifest('{"name": "${{MISSING}}"}', ctx, {
        envs: {},
        strict: false,
      });
      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value).to.contain("${{MISSING}}");
      }
    });

    it("should default to strict mode", async () => {
      const ctx = createMockContext();
      const result = await resolveManifest('{"name": "${{UNSET}}"}', ctx, { envs: {} });
      expect(result.isErr()).to.be.true;
    });

    it("should combine function expansion and env resolution", async () => {
      const dir = await setup();
      const dataFile = path.join(dir, "greeting.txt");
      await fs.writeFile(dataFile, "Hello ${{USER}}");

      const manifestPath = path.join(dir, "manifest.json");
      const ctx = createMockContext({ projectPath: dir });

      const result = await resolveManifest('{"msg": "$[file(\'greeting.txt\')]"}', ctx, {
        envs: { USER: "World" },
        manifestType: ManifestType.TeamsManifest,
        fromPath: manifestPath,
        strict: true,
      });
      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value).to.contain("Hello World");
      }
    });
  });
});

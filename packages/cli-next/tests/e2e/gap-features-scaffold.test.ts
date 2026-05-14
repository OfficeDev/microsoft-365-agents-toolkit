// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Scaffold-only E2E tests for templates that cannot run the full lifecycle
 * in CI today:
 *
 * - foundry-agent-to-m365: needs a real Foundry endpoint + Azure App Service
 * - da-meta-os-upgrade: gated behind the DAMetaOS feature flag
 *
 * These verify CLI argument wiring and output project structure via
 * subprocess invocation. No Azure resources or M365 tokens required.
 *
 * Once the blockers are resolved, these templates will be tested through
 * the standard lifecycle matrix in lifecycle.test.ts.
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

function tmpDir(): string {
  const dir = path.join(os.tmpdir(), `atk-gap-test-${Date.now()}`);
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

describe("Scaffold-only E2E (templates excluded from lifecycle)", function () {
  this.timeout(5 * 60 * 1000); // 5 min total

  let dir: string;

  beforeEach(function () {
    dir = tmpDir();
  });

  afterEach(async function () {
    await fs.promises.rm(dir, { recursive: true, force: true }).catch(() => {});
  });

  // ---------------------------------------------------------------------------
  // foundry-agent-to-m365 (blocked: no Foundry endpoint in CI)
  // ---------------------------------------------------------------------------
  describe("foundry-agent-to-m365: scaffold without Azure/Foundry", function () {
    it("--help shows foundryEndpoint and foundryAgentId options", async function () {
      const result = await run(`${ATK_BIN} new ai foundry-to-m365 --help`);
      expect(result.exitCode).to.equal(0);
      expect(result.stdout).to.include("foundryEndpoint");
      expect(result.stdout).to.include("foundryAgentId");
    });

    it("scaffolds project with foundry endpoint and agent ID", async function () {
      const result = await run(
        [
          ATK_BIN,
          "new",
          "ai",
          "foundry-to-m365",
          "--name",
          "FoundryApp",
          "--foundryEndpoint",
          "https://my-proj.services.ai.azure.com",
          "--foundryAgentId",
          "agent-00000000-0000-0000-0000-000000000001",
          "--folder",
          dir,
          "--non-interactive",
        ].join(" "),
        dir
      );
      expect(result.exitCode, `stderr: ${result.stderr}`).to.equal(0);

      const projectPath = path.join(dir, "FoundryApp");
      expect(fs.existsSync(projectPath), "project dir should exist").to.be.true;

      // Verify core project files
      expect(fs.existsSync(path.join(projectPath, "m365agents.yml")), "m365agents.yml").to.be.true;
      expect(fs.existsSync(path.join(projectPath, "env", ".env.dev")), ".env.dev").to.be.true;

      // Verify appPackage
      const appPackage = path.join(projectPath, "appPackage");
      expect(fs.existsSync(path.join(appPackage, "manifest.json")), "manifest.json").to.be.true;
    });

    it("scaffolds project without optional foundry values", async function () {
      const result = await run(
        [
          ATK_BIN,
          "new",
          "ai",
          "foundry-to-m365",
          "--name",
          "FoundryBasic",
          "--folder",
          dir,
          "--non-interactive",
        ].join(" "),
        dir
      );
      expect(result.exitCode, `stderr: ${result.stderr}`).to.equal(0);

      const projectPath = path.join(dir, "FoundryBasic");
      expect(fs.existsSync(projectPath), "project dir should exist").to.be.true;
    });
  });

  // ---------------------------------------------------------------------------
  // da-meta-os-upgrade (blocked: DAMetaOS feature flag off in CI)
  // ---------------------------------------------------------------------------
  describe("da-meta-os-upgrade: scaffold and project structure", function () {
    it("--help shows officeAddinFolder option", async function () {
      const result = await run(`${ATK_BIN} new da metaos-upgrade --help`);
      expect(result.exitCode).to.equal(0);
      expect(result.stdout).to.include("officeAddinFolder");
    });

    it("scaffold without source folder succeeds with warning", async function () {
      const result = await run(
        [
          ATK_BIN,
          "new",
          "da",
          "metaos-upgrade",
          "--name",
          "MetaOSTest",
          "--folder",
          dir,
          "--non-interactive",
        ].join(" "),
        dir
      );

      // Without source folder, descriptor produces a scaffold with warnings
      // but should still exit 0.
      expect(result.exitCode, `stderr: ${result.stderr}`).to.equal(0);

      const projectPath = path.join(dir, "MetaOSTest");
      expect(fs.existsSync(projectPath), "project dir should exist").to.be.true;
    });

    it("scaffold with source Office Add-in folder copies and upgrades", async function () {
      // Create a fake Office Add-in source project
      const sourceDir = path.join(dir, "source-addin");
      const srcPkg = path.join(sourceDir, "appPackage");
      await fs.promises.mkdir(path.join(sourceDir, "src", "commands"), { recursive: true });
      await fs.promises.mkdir(srcPkg, { recursive: true });
      await fs.promises.mkdir(path.join(sourceDir, "env"), { recursive: true });

      const manifest = {
        id: "00000000-0000-0000-0000-000000000000",
        name: { short: "TestAddin" },
        extensions: [
          {
            runtimes: [
              {
                code: { script: "commands.js" },
                actions: [],
              },
            ],
          },
        ],
      };
      await fs.promises.writeFile(
        path.join(srcPkg, "manifest.json"),
        JSON.stringify(manifest, null, 2)
      );
      await fs.promises.writeFile(
        path.join(sourceDir, "src", "commands", "commands.ts"),
        "// empty\n"
      );
      await fs.promises.writeFile(
        path.join(sourceDir, "package.json"),
        JSON.stringify({ devDependencies: {} })
      );
      await fs.promises.writeFile(path.join(sourceDir, "index.ts"), "// app");

      const result = await run(
        [
          ATK_BIN,
          "new",
          "da",
          "metaos-upgrade",
          "--name",
          "UpgradedApp",
          `--officeAddinFolder`,
          `"${sourceDir}"`,
          "--folder",
          dir,
          "--non-interactive",
        ].join(" "),
        dir
      );
      expect(result.exitCode, `stderr: ${result.stderr}`).to.equal(0);

      const projectPath = path.join(dir, "UpgradedApp");
      expect(fs.existsSync(projectPath), "project dir should exist").to.be.true;

      // Verify source files were copied
      expect(fs.existsSync(path.join(projectPath, "index.ts")), "copied index.ts").to.be.true;

      // Verify DA extension
      const upgradedManifest = readJson(path.join(projectPath, "appPackage", "manifest.json"));
      expect(upgradedManifest.copilotAgents).to.exist;

      // Verify project ID was unified (not the original zero UUID)
      expect(upgradedManifest.id).to.not.equal("00000000-0000-0000-0000-000000000000");
    });
  });
});

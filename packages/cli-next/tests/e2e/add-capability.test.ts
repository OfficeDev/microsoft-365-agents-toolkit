// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Add-capability tests — subprocess layer.
 *
 * Verifies that `atk add capability` modifies an existing Declarative Agent
 * project correctly (e.g., adding web-search knowledge). Each test scaffolds
 * a fresh DA project, then runs `atk add capability` on it.
 *
 * Does NOT create real Azure resources.
 *
 * Maps to ADO test plan suites "DA - Embedded Knowledge" (34657654) and
 * old E2E tests declarativeAgent/addKnowledge/.
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
  const dir = path.join(os.tmpdir(), `atk-addcap-test-${Date.now()}`);
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

async function scaffoldDA(dir: string, name: string): Promise<string> {
  const result = await run(
    `${ATK_BIN} new da basic --name ${name} --folder ${dir} --non-interactive`,
    dir
  );
  if (result.exitCode !== 0) {
    throw new Error(`Scaffold failed: ${result.stderr}`);
  }
  return path.join(dir, name);
}

describe("Add capability tests", function () {
  this.timeout(5 * 60 * 1000); // 5 min total

  let dir: string;

  beforeEach(function () {
    dir = tmpDir();
  });

  afterEach(async function () {
    await fs.promises.rm(dir, { recursive: true, force: true }).catch(() => {});
  });

  describe("atk add capability --help", function () {
    it("exits 0 and shows --capability-type", async function () {
      const result = await run(`${ATK_BIN} add capability --help`);
      expect(result.exitCode).to.equal(0);
      expect(result.stdout).to.include("--capability-type");
    });

    it("shows supported capability types", async function () {
      const result = await run(`${ATK_BIN} add capability --help`);
      expect(result.exitCode).to.equal(0);
      expect(result.stdout).to.include("web-search");
    });
  });

  describe("add web-search capability to DA project", function () {
    it("atk add capability --capability-type web-search succeeds", async function () {
      const projectPath = await scaffoldDA(dir, "AddCapDA");

      const result = await run(
        `${ATK_BIN} add capability --capability-type web-search`,
        projectPath
      );
      expect(result.exitCode, `stderr: ${result.stderr}`).to.equal(0);

      // Verify declarativeAgent.json was updated with web_search capability
      const daManifestPath = path.join(projectPath, "appPackage", "declarativeAgent.json");
      expect(fs.existsSync(daManifestPath), "declarativeAgent.json should exist").to.be.true;
      const daManifest = readJson(daManifestPath);
      expect(daManifest.capabilities).to.be.an("array");
      const capabilities = daManifest.capabilities as Record<string, unknown>[];
      const webSearchCap = capabilities.find(
        (c) => c.name === "WebSearch" || c.name === "web_search"
      );
      expect(webSearchCap, "should have web_search capability").to.not.be.undefined;
    });
  });

  describe("add capability error cases", function () {
    it("fails without --capability-type", async function () {
      const projectPath = await scaffoldDA(dir, "ErrCapDA");

      const result = await run(`${ATK_BIN} add capability`, projectPath);
      expect(result.exitCode).to.not.equal(0);
      expect(result.stderr).to.match(/required|missing|capability-type/i);
    });

    it("fails with invalid --capability-type", async function () {
      const projectPath = await scaffoldDA(dir, "InvalidCapDA");

      const result = await run(
        `${ATK_BIN} add capability --capability-type invalid-type`,
        projectPath
      );
      expect(result.exitCode).to.not.equal(0);
    });
  });
});

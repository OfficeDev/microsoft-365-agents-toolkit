// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * CLI syntax verification tests — subprocess layer.
 *
 * Verifies that the CLI binary parses arguments correctly and produces
 * expected exit codes. Does NOT create real Azure resources.
 * Uses ATK_BIN env var for workspace vs registry switching.
 *
 * Tests:
 *   - `atk new <category> <slug> --name X --non-interactive` → scaffold only
 *   - `atk --version` → exit 0
 *   - `atk provision --help` → exit 0
 *   - Invalid command → non-zero exit
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
  const dir = path.join(os.tmpdir(), `atk-cli-test-${Date.now()}`);
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

describe("CLI syntax verification", function () {
  this.timeout(3 * 60 * 1000); // 3 min total

  it("atk --version exits 0", async function () {
    const result = await run(`${ATK_BIN} --version`);
    expect(result.exitCode).to.equal(0);
    expect(result.stdout).to.match(/\d+\.\d+\.\d+/);
  });

  it("atk provision --help exits 0", async function () {
    const result = await run(`${ATK_BIN} provision --help`);
    expect(result.exitCode).to.equal(0);
    expect(result.stdout).to.include("provision");
  });

  it("atk invalid-command exits non-zero", async function () {
    const result = await run(`${ATK_BIN} nonexistent-command-xyz`);
    expect(result.exitCode).to.not.equal(0);
  });

  describe("atk new (scaffold only)", function () {
    let dir: string;

    beforeEach(function () {
      dir = tmpDir();
    });

    afterEach(async function () {
      await fs.promises.rm(dir, { recursive: true, force: true }).catch(() => {});
    });

    it("atk new bot echo --name TestBot --non-interactive", async function () {
      const result = await run(
        `${ATK_BIN} new bot echo --name TestBot --folder ${dir} --non-interactive`,
        dir
      );
      // The command may output to stderr for debug info but should exit 0
      expect(result.exitCode, `stderr: ${result.stderr}`).to.equal(0);
      const projectPath = path.join(dir, "TestBot");
      expect(fs.existsSync(projectPath), `project dir should exist at ${projectPath}`).to.be.true;
    });

    it("atk new da basic --name TestDA --non-interactive", async function () {
      const result = await run(
        `${ATK_BIN} new da basic --name TestDA --folder ${dir} --non-interactive`,
        dir
      );
      expect(result.exitCode, `stderr: ${result.stderr}`).to.equal(0);
      const projectPath = path.join(dir, "TestDA");
      expect(fs.existsSync(projectPath), `project dir should exist`).to.be.true;
    });

    it("atk new tab basic --name TestTab --non-interactive", async function () {
      const result = await run(
        `${ATK_BIN} new tab basic --name TestTab --folder ${dir} --non-interactive`,
        dir
      );
      expect(result.exitCode, `stderr: ${result.stderr}`).to.equal(0);
      const projectPath = path.join(dir, "TestTab");
      expect(fs.existsSync(projectPath), `project dir should exist`).to.be.true;
    });

    it("atk new cea basic --name TestCEA --non-interactive", async function () {
      const result = await run(
        `${ATK_BIN} new cea basic --name TestCEA --folder ${dir} --non-interactive`,
        dir
      );
      expect(result.exitCode, `stderr: ${result.stderr}`).to.equal(0);
      const projectPath = path.join(dir, "TestCEA");
      expect(fs.existsSync(projectPath), `project dir should exist`).to.be.true;
    });

    it("atk new me search-based --name TestME --non-interactive", async function () {
      const result = await run(
        `${ATK_BIN} new me search-based --name TestME --folder ${dir} --non-interactive`,
        dir
      );
      expect(result.exitCode, `stderr: ${result.stderr}`).to.equal(0);
      const projectPath = path.join(dir, "TestME");
      expect(fs.existsSync(projectPath), `project dir should exist`).to.be.true;
    });
  });

  describe("Add / Set / M365 command --help", function () {
    it("atk add --help exits 0 and shows subcommands", async function () {
      const result = await run(`${ATK_BIN} add --help`);
      expect(result.exitCode).to.equal(0);
      expect(result.stdout).to.include("action");
      expect(result.stdout).to.include("capability");
      expect(result.stdout).to.include("auth-config");
    });

    it("atk add action --help exits 0 and shows required options", async function () {
      const result = await run(`${ATK_BIN} add action --help`);
      expect(result.exitCode).to.equal(0);
      expect(result.stdout).to.include("--api-spec-path");
      expect(result.stdout).to.include("--plugin-manifest-path");
      expect(result.stdout).to.include("--action-id");
    });

    it("atk add capability --help exits 0 and shows --capability-type", async function () {
      const result = await run(`${ATK_BIN} add capability --help`);
      expect(result.exitCode).to.equal(0);
      expect(result.stdout).to.include("--capability-type");
    });

    it("atk set sensitivityLabel --help exits 0 and shows --label-id", async function () {
      const result = await run(`${ATK_BIN} set sensitivityLabel --help`);
      expect(result.exitCode).to.equal(0);
      expect(result.stdout).to.include("--label-id");
    });

    it("atk m365-sideload --help exits 0 and shows options", async function () {
      const result = await run(`${ATK_BIN} m365-sideload --help`);
      expect(result.exitCode).to.equal(0);
      expect(result.stdout).to.include("--file-path");
      expect(result.stdout).to.include("--scope");
    });
  });
});

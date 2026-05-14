// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Auth command tests — subprocess layer.
 *
 * Verifies that `atk auth` subcommands (show, login, logout) parse
 * arguments correctly and produce expected help output. Does NOT perform
 * actual authentication flows (those require interactive credentials).
 *
 * Maps to ADO test plan suites "[CLI] command" (24569138) and
 * "[CLI] help" (24857098).
 */

import { expect } from "chai";
import { describe, it } from "mocha";
import { promisify } from "util";
import { exec } from "child_process";

const execAsync = promisify(exec);

const ATK_BIN = process.env.ATK_BIN || "atk";
const TIMEOUT = 60_000; // 1 min per command

async function run(command: string): Promise<{ stdout: string; stderr: string; exitCode: number }> {
  try {
    const { stdout, stderr } = await execAsync(command, {
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

describe("Auth command verification", function () {
  this.timeout(3 * 60 * 1000); // 3 min total

  describe("atk auth --help", function () {
    it("exits 0 and shows auth subcommands", async function () {
      const result = await run(`${ATK_BIN} auth --help`);
      expect(result.exitCode).to.equal(0);
      expect(result.stdout).to.include("show");
      expect(result.stdout).to.include("login");
      expect(result.stdout).to.include("logout");
    });
  });

  describe("atk auth show", function () {
    it("atk auth show --help exits 0", async function () {
      const result = await run(`${ATK_BIN} auth show --help`);
      expect(result.exitCode).to.equal(0);
    });
  });

  describe("atk auth login", function () {
    it("atk auth login --help exits 0 and shows azure/m365 subcommands", async function () {
      const result = await run(`${ATK_BIN} auth login --help`);
      expect(result.exitCode).to.equal(0);
      expect(result.stdout).to.include("azure");
      expect(result.stdout).to.include("m365");
    });

    it("atk auth login azure --help shows --tenant and --service-principal", async function () {
      const result = await run(`${ATK_BIN} auth login azure --help`);
      expect(result.exitCode).to.equal(0);
      expect(result.stdout).to.include("--tenant");
      expect(result.stdout).to.include("--service-principal");
    });

    it("atk auth login m365 --help shows --tenant", async function () {
      const result = await run(`${ATK_BIN} auth login m365 --help`);
      expect(result.exitCode).to.equal(0);
      expect(result.stdout).to.include("--tenant");
    });
  });

  describe("atk auth logout", function () {
    it("atk auth logout --help exits 0", async function () {
      const result = await run(`${ATK_BIN} auth logout --help`);
      expect(result.exitCode).to.equal(0);
    });
  });
});

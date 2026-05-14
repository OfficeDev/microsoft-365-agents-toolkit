/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import { expect } from "chai";
import { describe, it, afterEach, beforeEach } from "mocha";
import * as sinon from "sinon";
import { buildProgram } from "../../src/commands";
import { cliTelemetry } from "../../src/telemetry";

/**
 * Integration tests that exercise core-next operations through the CLI.
 *
 * These tests mock only I/O boundaries (telemetry, console).
 * Core-next operations (env management, template registry, validation) run for real.
 */
describe("Core-next Integration", () => {
  const sandbox = sinon.createSandbox();
  let tmpDir: string;

  beforeEach(async () => {
    sandbox.stub(cliTelemetry, "sendEvent");
    sandbox.stub(cliTelemetry, "sendErrorEvent");
    sandbox.stub(cliTelemetry, "flush").resolves();
    sandbox.stub(console, "log");
    sandbox.stub(console, "warn");
    sandbox.stub(console, "error");
    tmpDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), "cli-integ-"));
  });

  afterEach(async () => {
    sandbox.restore();
    process.exitCode = undefined;
    await fs.promises.rm(tmpDir, { recursive: true, force: true });
  });

  describe("Environment commands (real core-next)", () => {
    it("atk env list — should list environments from project dir", async () => {
      // Setup: create env directory with .env files
      const envDir = path.join(tmpDir, "env");
      await fs.promises.mkdir(envDir, { recursive: true });
      await fs.promises.writeFile(path.join(envDir, ".env.dev"), "APP_NAME=test\n");
      await fs.promises.writeFile(path.join(envDir, ".env.staging"), "APP_NAME=test\n");

      const program = buildProgram("atk");
      program.exitOverride();

      await program.parseAsync(["node", "atk", "env", "list", "--project-folder", tmpDir]);

      // Verify console.log was called with env names
      const logStub = console.log as sinon.SinonStub;
      const output = logStub
        .getCalls()
        .map((c) => c.args.join(" "))
        .join("\n");
      expect(output).to.include("dev");
      expect(output).to.include("staging");
    });

    it("atk env add — should create a new environment file", async () => {
      const program = buildProgram("atk");
      program.exitOverride();

      await program.parseAsync([
        "node",
        "atk",
        "env",
        "add",
        "production",
        "--project-folder",
        tmpDir,
      ]);

      const envFilePath = path.join(tmpDir, "env", ".env.production");
      const exists = await fileExists(envFilePath);
      expect(exists).to.be.true;
    });

    it("atk env reset — should clear environment file", async () => {
      const envDir = path.join(tmpDir, "env");
      await fs.promises.mkdir(envDir, { recursive: true });
      await fs.promises.writeFile(path.join(envDir, ".env.dev"), "APP_NAME=myapp\nSECRET=abc\n");

      const program = buildProgram("atk");
      program.exitOverride();

      await program.parseAsync(["node", "atk", "env", "reset", "dev", "--project-folder", tmpDir]);

      const content = await fs.promises.readFile(path.join(envDir, ".env.dev"), "utf-8");
      expect(content.trim()).to.equal("");
    });
  });

  describe("List commands (real core-next)", () => {
    it("atk list templates — should display registered templates", async () => {
      const program = buildProgram("atk");
      program.exitOverride();

      await program.parseAsync(["node", "atk", "list", "templates"]);

      // The list command calls registerBuiltinTemplates and prints them.
      const logStub = console.log as sinon.SinonStub;
      const output = logStub
        .getCalls()
        .map((c) => c.args.join(" "))
        .join("\n");
      // Should contain at least one known template
      expect(output).to.include("bot");
    });
  });

  describe("New command tree (real core-next registry)", () => {
    it("atk new --help should show category subcommands", () => {
      const program = buildProgram("atk");
      program.exitOverride();

      const newCmd = program.commands.find((c) => c.name() === "new");
      expect(newCmd).to.exist;

      // Should have category subcommands from the registry
      const subcmdNames = newCmd!.commands.map((c) => c.name());
      expect(subcmdNames).to.include("da");
      expect(subcmdNames).to.include("bot");
    });

    it("atk new da --help should show template subcommands", () => {
      const program = buildProgram("atk");
      program.exitOverride();

      const newCmd = program.commands.find((c) => c.name() === "new");
      const daCmd = newCmd!.commands.find((c) => c.name() === "da");
      expect(daCmd).to.exist;
      expect(daCmd!.commands.length).to.be.greaterThan(0);

      // Should contain at least "basic"
      const templateNames = daCmd!.commands.map((c) => c.name());
      expect(templateNames).to.include("basic");
    });

    it("atk new bot --help should show bot template subcommands", () => {
      const program = buildProgram("atk");
      program.exitOverride();

      const newCmd = program.commands.find((c) => c.name() === "new");
      const botCmd = newCmd!.commands.find((c) => c.name() === "bot");
      expect(botCmd).to.exist;
      expect(botCmd!.commands.length).to.be.greaterThan(0);

      // Should contain at least "echo"
      const templateNames = botCmd!.commands.map((c) => c.name());
      expect(templateNames).to.include("echo");
    });

    it("bot/echo subcommand should have --name (required) and --folder options", () => {
      const program = buildProgram("atk");
      program.exitOverride();

      const newCmd = program.commands.find((c) => c.name() === "new");
      const botCmd = newCmd!.commands.find((c) => c.name() === "bot");
      const echoCmd = botCmd!.commands.find((c) => c.name() === "echo");
      expect(echoCmd).to.exist;

      const optNames = echoCmd!.options.map((o) => o.long);
      expect(optNames).to.include("--name");
      expect(optNames).to.include("--folder");
    });
  });
});

async function fileExists(filePath: string): Promise<boolean> {
  try {
    await fs.promises.access(filePath);
    return true;
  } catch {
    return false;
  }
}

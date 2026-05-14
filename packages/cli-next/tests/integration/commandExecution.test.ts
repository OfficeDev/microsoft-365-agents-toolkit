/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, afterEach, beforeEach } from "mocha";
import * as sinon from "sinon";
import { buildProgram } from "../../src/commands";
import { cliTelemetry } from "../../src/telemetry";

describe("CLI Integration", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    sandbox.stub(cliTelemetry, "sendEvent");
    sandbox.stub(cliTelemetry, "sendErrorEvent");
    sandbox.stub(cliTelemetry, "flush").resolves();
    sandbox.stub(console, "log");
    sandbox.stub(console, "error");
  });

  afterEach(() => {
    sandbox.restore();
    process.exitCode = undefined;
  });

  describe("end-to-end program parse", () => {
    it("should parse --version without throwing", () => {
      const program = buildProgram("atk");
      program.exitOverride();
      try {
        program.parse(["node", "atk", "--version"]);
      } catch (e: unknown) {
        // commander throws a special CommanderError for --version
        if (e instanceof Error && "exitCode" in e) {
          expect((e as { exitCode: number }).exitCode).to.equal(0);
        }
      }
    });

    it("should have all expected top-level commands", () => {
      const program = buildProgram("atk");
      const names = program.commands.map((c) => c.name());
      // Verify all command groups from commands/index.ts
      for (const expected of ["new", "auth", "env", "add", "list"]) {
        expect(names).to.include(expected);
      }
    });

    it("should parse global options without throwing", () => {
      const program = buildProgram("atk");
      program.exitOverride();
      try {
        program.parse(["node", "atk", "--output", "json", "--non-interactive", "--help"], {
          from: "user",
        });
      } catch (e: unknown) {
        // --help triggers CommanderError with code 'commander.helpDisplayed'
        if (e instanceof Error && "exitCode" in e) {
          expect((e as { exitCode: number }).exitCode).to.equal(0);
        }
      }
    });
  });

  describe("error integration", () => {
    it("should import error classes from the same package", () => {
      const {
        MissingRequiredOptionError,
        InvalidChoiceError,
        CLISystemError,
      } = require("../../src/error");
      const e1 = new MissingRequiredOptionError("cmd", "opt");
      expect(e1.source).to.equal("CLI");
      const e2 = new InvalidChoiceError("opt", "bad", ["a", "b"]);
      expect(e2.source).to.equal("CLI");
      const e3 = new CLISystemError("fail");
      expect(e3.source).to.equal("CLI");
    });
  });
});

/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it } from "mocha";
import { buildProgram } from "../../src/commands";

describe("CLI Commands", () => {
  describe("buildProgram()", () => {
    it("should create a commander program with subcommands", () => {
      const program = buildProgram("atk");
      expect(program.name()).to.equal("atk");
      expect(program.commands.length).to.be.greaterThan(0);
    });

    it("should include core command groups", () => {
      const program = buildProgram("atk");
      const names = program.commands.map((c) => c.name());
      expect(names).to.include("new");
      expect(names).to.include("auth");
      expect(names).to.include("env");
    });

    it("should have --output and --non-interactive global options", () => {
      const program = buildProgram("atk");
      const optionNames = program.options.map((o) => o.long);
      expect(optionNames).to.include("--output");
      expect(optionNames).to.include("--non-interactive");
    });

    it("should have version flag", () => {
      const program = buildProgram("atk");
      const optionNames = program.options.map((o) => o.long);
      expect(optionNames).to.include("--version");
    });
  });
});

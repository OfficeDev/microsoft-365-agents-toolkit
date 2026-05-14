/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it } from "mocha";
import {
  MissingRequiredOptionError,
  MissingRequiredArgumentError,
  InvalidChoiceError,
  UnknownCommandError,
  CLISystemError,
} from "../../src/error";

describe("CLI Error types", () => {
  describe("MissingRequiredOptionError", () => {
    it("should include command and option names in message", () => {
      const err = new MissingRequiredOptionError("create", "name");
      expect(err.message).to.include("--name");
      expect(err.message).to.include("create");
      expect(err.name).to.equal("MissingRequiredOptionError");
    });
  });

  describe("MissingRequiredArgumentError", () => {
    it("should include command and argument names", () => {
      const err = new MissingRequiredArgumentError("deploy", "project-path");
      expect(err.message).to.include("<project-path>");
      expect(err.message).to.include("deploy");
    });
  });

  describe("InvalidChoiceError", () => {
    it("should list valid choices", () => {
      const err = new InvalidChoiceError("output", "xml", ["text", "json"]);
      expect(err.message).to.include("xml");
      expect(err.message).to.include("text, json");
    });
  });

  describe("UnknownCommandError", () => {
    it("should show suggestion when provided", () => {
      const err = new UnknownCommandError("ceate", "create");
      expect(err.message).to.include("Did you mean");
      expect(err.message).to.include("create");
    });

    it("should show base message without suggestion", () => {
      const err = new UnknownCommandError("xyz");
      expect(err.message).to.include("Unknown command");
      expect(err.message).to.include("xyz");
    });
  });

  describe("CLISystemError", () => {
    it("should wrap inner error", () => {
      const inner = new Error("network failure");
      const err = new CLISystemError("Service call failed", inner);
      expect(err.message).to.include("Service call failed");
      expect(err.name).to.equal("CLISystemError");
    });

    it("should work without inner error", () => {
      const err = new CLISystemError("Something went wrong");
      expect(err.message).to.equal("Something went wrong");
    });
  });
});

/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it } from "mocha";
import { userError, systemError } from "../../../src/core/error";

describe("AtkError", () => {
  describe("userError()", () => {
    it("should create a user error with required fields", () => {
      const err = userError("BadInput", "Input is invalid");
      expect(err.code).to.equal("BadInput");
      expect(err.message).to.equal("Input is invalid");
      expect(err.kind).to.equal("user");
      expect(err.source).to.be.undefined;
    });

    it("should include optional fields", () => {
      const inner = new Error("root cause");
      const err = userError("MissingFile", "File not found", {
        source: "scaffold",
        help: "https://docs.example.com",
        inner,
        displayMessage: "Please check the file path",
      });
      expect(err.source).to.equal("scaffold");
      expect(err.help).to.equal("https://docs.example.com");
      expect(err.inner).to.equal(inner);
      expect(err.displayMessage).to.equal("Please check the file path");
    });
  });

  describe("systemError()", () => {
    it("should create a system error", () => {
      const err = systemError("ServiceDown", "API returned 500");
      expect(err.code).to.equal("ServiceDown");
      expect(err.kind).to.equal("system");
    });
  });
});

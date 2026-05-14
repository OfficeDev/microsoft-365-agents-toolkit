/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it } from "mocha";
import { anonymizeFilePaths, sanitizeProperties } from "../../src/telemetry/sanitize";

describe("sanitize", () => {
  describe("anonymizeFilePaths()", () => {
    it("should return empty string for undefined/null/empty", () => {
      expect(anonymizeFilePaths(undefined)).to.equal("");
      expect(anonymizeFilePaths("")).to.equal("");
    });

    it("should redact absolute file paths", () => {
      const input = "Error at C:\\Users\\alice\\project\\src\\index.ts:10";
      const result = anonymizeFilePaths(input);
      expect(result).to.include("<REDACTED: user-file-path>");
      expect(result).not.to.include("alice");
    });

    it("should redact unix file paths", () => {
      const input = "Failed at /home/alice/project/src/index.ts";
      const result = anonymizeFilePaths(input);
      expect(result).to.include("<REDACTED: user-file-path>");
      expect(result).not.to.include("alice");
    });

    it("should preserve node_modules paths", () => {
      const input = "node_modules/commander/index.js threw an error";
      const result = anonymizeFilePaths(input);
      expect(result).to.include("node_modules/commander");
    });

    it("should not change strings without file paths", () => {
      const input = "simple error message";
      expect(anonymizeFilePaths(input)).to.equal("simple error message");
    });
  });

  describe("sanitizeProperties()", () => {
    it("should return undefined for undefined input", () => {
      expect(sanitizeProperties(undefined)).to.be.undefined;
    });

    it("should redact properties containing tokens", () => {
      const props = { auth: "bearer token=abc123xyz" };
      const cleaned = sanitizeProperties(props)!;
      expect(cleaned.auth).to.equal("<REDACTED: token>");
    });

    it("should redact properties containing passwords", () => {
      const props = { config: "password=secret123" };
      const cleaned = sanitizeProperties(props)!;
      expect(cleaned.config).to.equal("<REDACTED: password>");
    });

    it("should redact email addresses", () => {
      const props = { user: "user@example.com" };
      const cleaned = sanitizeProperties(props)!;
      expect(cleaned.user).to.equal("<REDACTED: email>");
    });

    it("should pass through safe values", () => {
      const props = { command: "provision", success: "true" };
      const cleaned = sanitizeProperties(props)!;
      expect(cleaned.command).to.equal("provision");
      expect(cleaned.success).to.equal("true");
    });

    it("should anonymize file paths in values", () => {
      const props = { stack: "Error at C:\\Users\\bob\\code\\app.ts" };
      const cleaned = sanitizeProperties(props)!;
      expect(cleaned.stack).to.include("<REDACTED: user-file-path>");
      expect(cleaned.stack).not.to.include("bob");
    });
  });
});

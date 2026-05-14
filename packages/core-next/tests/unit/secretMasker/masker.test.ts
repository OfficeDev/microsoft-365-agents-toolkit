/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it } from "mocha";
import { CREDENTIAL_KEYWORDS, matchesCredentialKeyword } from "../../../src/secretMasker/keywords";
import { maskSecret, maskSecretValues } from "../../../src/secretMasker/masker";

describe("Secret Masker", () => {
  describe("CREDENTIAL_KEYWORDS", () => {
    it("should have a comprehensive set of credential keywords", () => {
      expect(CREDENTIAL_KEYWORDS.length).to.be.greaterThanOrEqual(50);
    });

    it("should include common credential suffixes", () => {
      expect(CREDENTIAL_KEYWORDS).to.include("password");
      expect(CREDENTIAL_KEYWORDS).to.include("secret");
      expect(CREDENTIAL_KEYWORDS).to.include("apikey");
      expect(CREDENTIAL_KEYWORDS).to.include("connectionstring");
    });
  });

  describe("matchesCredentialKeyword()", () => {
    it("should match keys ending with credential keywords", () => {
      expect(matchesCredentialKeyword("MY_PASSWORD")).to.be.true;
      expect(matchesCredentialKeyword("db-connection-string")).to.be.false;
      expect(matchesCredentialKeyword("dbconnectionstring")).to.be.true;
      expect(matchesCredentialKeyword("SECRET")).to.be.true;
      expect(matchesCredentialKeyword("MY_API_KEY")).to.be.false; // "key" alone is not a suffix
    });

    it("should be case-insensitive", () => {
      expect(matchesCredentialKeyword("myPassword")).to.be.true;
      expect(matchesCredentialKeyword("MYSECRET")).to.be.true;
    });

    it("should not match non-credential keys", () => {
      expect(matchesCredentialKeyword("username")).to.be.false;
      expect(matchesCredentialKeyword("hostname")).to.be.false;
      expect(matchesCredentialKeyword("PORT")).to.be.false;
    });
  });

  describe("maskSecret()", () => {
    it("should mask values for credential keys in key=value patterns", () => {
      const input = "password=hunter2&username=alice";
      const result = maskSecret(input);
      expect(result).to.include("password=***");
      expect(result).to.include("username=alice");
    });

    it("should handle key:value patterns", () => {
      const input = "password: hunter2";
      const result = maskSecret(input);
      expect(result).to.include("password: ***");
    });

    it("should use custom replacement string", () => {
      const input = "password=hunter2";
      const result = maskSecret(input, "[REDACTED]");
      expect(result).to.include("[REDACTED]");
    });

    it("should return empty/falsy input unchanged", () => {
      expect(maskSecret("")).to.equal("");
    });

    it("should not mask non-credential keys", () => {
      const input = "hostname=myserver&port=3000";
      expect(maskSecret(input)).to.equal("hostname=myserver&port=3000");
    });
  });

  describe("maskSecretValues()", () => {
    it("should mask values of credential-matching keys in a record", () => {
      const record = {
        MY_PASSWORD: "secret123",
        MY_HOST: "localhost",
        API_ACCESSTOKEN: "token-abc",
      };
      const result = maskSecretValues(record);
      expect(result.MY_PASSWORD).to.equal("***");
      expect(result.MY_HOST).to.equal("localhost");
      expect(result.API_ACCESSTOKEN).to.equal("***");
    });

    it("should not mutate the original record", () => {
      const record = { MY_PASSWORD: "secret123" };
      const result = maskSecretValues(record);
      expect(record.MY_PASSWORD).to.equal("secret123");
      expect(result.MY_PASSWORD).to.equal("***");
    });
  });
});

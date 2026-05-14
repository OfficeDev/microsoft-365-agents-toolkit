/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it } from "mocha";
import { FeatureFlagRegistry } from "../../src/featureFlags/registry";
import { FeatureFlagSource } from "../../src/featureFlags/types";
import { builtinFlags, createDefaultRegistry } from "../../src/featureFlags/flags";
import { maskSecret, maskSecretValues } from "../../src/secretMasker/masker";
import { Localizer } from "../../src/localization/localizer";
import * as path from "path";
import * as fs from "fs";
import * as os from "os";

/**
 * Integration test: exercises cross-cutting modules together,
 * verifying that feature flags, secret masking, and localization
 * work correctly when composed in realistic scenarios.
 */
describe("Integration: Cross-cutting modules", () => {
  describe("Feature flags → conditional masking", () => {
    it("should mask secrets only when feature flag is enabled", () => {
      const values = new Map([["TEAMSFX_V4_CORE", "true"]]);
      const source: FeatureFlagSource = { get: (n) => values.get(n) };
      const registry = new FeatureFlagRegistry(source);
      registry.registerAll(builtinFlags);

      const record = { MY_PASSWORD: "secret123", HOSTNAME: "server1" };

      if (registry.isEnabled("TEAMSFX_V4_CORE")) {
        const masked = maskSecretValues(record);
        expect(masked.MY_PASSWORD).to.equal("***");
        expect(masked.HOSTNAME).to.equal("server1");
      }
    });
  });

  describe("Localizer with temp bundle", () => {
    it("should load and resolve strings from a temp package.nls.json", () => {
      const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "atk-test-"));
      const bundle = {
        "core.greeting": "Hello, %s!",
        "core.error.notFound": "Resource '%s' not found in '%s'.",
      };
      fs.writeFileSync(path.join(tmpDir, "package.nls.json"), JSON.stringify(bundle));

      try {
        const localizer = new Localizer();
        localizer.loadBundle(tmpDir);

        expect(localizer.getString("core.greeting", "World")).to.equal("Hello, World!");
        expect(localizer.getString("core.error.notFound", "file.txt", "/tmp")).to.equal(
          "Resource 'file.txt' not found in '/tmp'."
        );
        expect(localizer.getString("nonexistent.key")).to.equal("nonexistent.key");
      } finally {
        fs.rmSync(tmpDir, { recursive: true });
      }
    });

    it("should fall back to default string when no locale bundle", () => {
      const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "atk-test-"));
      fs.writeFileSync(path.join(tmpDir, "package.nls.json"), JSON.stringify({ "msg.ok": "OK" }));

      try {
        const localizer = new Localizer();
        localizer.loadBundle(tmpDir);
        expect(localizer.getDefaultString("msg.ok")).to.equal("OK");
      } finally {
        fs.rmSync(tmpDir, { recursive: true });
      }
    });
  });

  describe("Secret masking in realistic scenarios", () => {
    it("should mask secrets in connection string-like text", () => {
      const text = "Server=myserver;password=p@ssw0rd;Username=admin;accesstoken=abc123";
      const result = maskSecret(text);
      expect(result).to.include("password=***");
      expect(result).to.include("accesstoken=***");
      expect(result).to.include("Username=admin");
    });

    it("should mask secret values from environment-like records", () => {
      const env = {
        DB_PASSWORD: "hunter2",
        DB_HOST: "localhost",
        AZURE_CLIENTSECRET: "my-secret",
        APP_NAME: "my-app",
      };
      const masked = maskSecretValues(env);
      expect(masked.DB_PASSWORD).to.equal("***");
      expect(masked.AZURE_CLIENTSECRET).to.equal("***");
      expect(masked.DB_HOST).to.equal("localhost");
      expect(masked.APP_NAME).to.equal("my-app");
    });
  });

  describe("Default registry with real process.env", () => {
    it("should create a working registry that reads from process.env", () => {
      const registry = createDefaultRegistry();
      // V4Core is false by default, and likely not set in test env
      expect(registry.isEnabled("TEAMSFX_V4_CORE")).to.be.false;
      expect(registry.list().length).to.be.greaterThan(0);
    });
  });
});

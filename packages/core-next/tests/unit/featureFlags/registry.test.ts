/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it } from "mocha";
import { FeatureFlagRegistry } from "../../../src/featureFlags/registry";
import { FeatureFlag, FeatureFlagSource } from "../../../src/featureFlags/types";
import { builtinFlags, createDefaultRegistry } from "../../../src/featureFlags/flags";

describe("Feature Flags", () => {
  describe("FeatureFlagRegistry", () => {
    it("should register and retrieve flags", () => {
      const registry = new FeatureFlagRegistry();
      const flag: FeatureFlag = { name: "TEST_FLAG", defaultValue: false };
      registry.register(flag);
      expect(registry.list()).to.have.length(1);
      expect(registry.list()[0].name).to.equal("TEST_FLAG");
    });

    it("should return false for unregistered flags", () => {
      const registry = new FeatureFlagRegistry();
      expect(registry.isEnabled("NONEXISTENT")).to.be.false;
    });

    it("should use default value when env var is not set", () => {
      const source: FeatureFlagSource = { get: () => undefined };
      const registry = new FeatureFlagRegistry(source);
      registry.register({ name: "FLAG_A", defaultValue: true });
      registry.register({ name: "FLAG_B", defaultValue: false });
      expect(registry.isEnabled("FLAG_A")).to.be.true;
      expect(registry.isEnabled("FLAG_B")).to.be.false;
    });

    it("should use env var value when set to 'true'", () => {
      const source: FeatureFlagSource = { get: () => "true" };
      const registry = new FeatureFlagRegistry(source);
      registry.register({ name: "FLAG", defaultValue: false });
      expect(registry.isEnabled("FLAG")).to.be.true;
    });

    it("should treat '1' as enabled", () => {
      const source: FeatureFlagSource = { get: () => "1" };
      const registry = new FeatureFlagRegistry(source);
      registry.register({ name: "FLAG", defaultValue: false });
      expect(registry.isEnabled("FLAG")).to.be.true;
    });

    it("should treat 'false' as disabled", () => {
      const source: FeatureFlagSource = { get: () => "false" };
      const registry = new FeatureFlagRegistry(source);
      registry.register({ name: "FLAG", defaultValue: true });
      expect(registry.isEnabled("FLAG")).to.be.false;
    });

    it("should treat empty string as unset (use default)", () => {
      const source: FeatureFlagSource = { get: () => "" };
      const registry = new FeatureFlagRegistry(source);
      registry.register({ name: "FLAG", defaultValue: true });
      expect(registry.isEnabled("FLAG")).to.be.true;
    });

    it("registerAll should register multiple flags", () => {
      const registry = new FeatureFlagRegistry();
      registry.registerAll([
        { name: "A", defaultValue: false },
        { name: "B", defaultValue: true },
      ]);
      expect(registry.list()).to.have.length(2);
    });

    it("listEnabled should return only enabled flags", () => {
      const values = new Map([
        ["A", "true"],
        ["B", "false"],
      ]);
      const source: FeatureFlagSource = { get: (n) => values.get(n) };
      const registry = new FeatureFlagRegistry(source);
      registry.registerAll([
        { name: "A", defaultValue: false },
        { name: "B", defaultValue: false },
      ]);
      const enabled = registry.listEnabled();
      expect(enabled).to.have.length(1);
      expect(enabled[0].name).to.equal("A");
    });

    it("getValue should return raw env var value", () => {
      const source: FeatureFlagSource = { get: () => "custom-value" };
      const registry = new FeatureFlagRegistry(source);
      expect(registry.getValue("ANY")).to.equal("custom-value");
    });
  });

  describe("builtinFlags", () => {
    it("should include V4Core flag", () => {
      const v4 = builtinFlags.find((f) => f.name === "TEAMSFX_V4_CORE");
      expect(v4).to.exist;
      expect(v4!.defaultValue).to.be.false;
    });

    it("should include MCPForDA flag with default true", () => {
      const mcp = builtinFlags.find((f) => f.name === "TEAMSFX_MCP_FOR_DA");
      expect(mcp).to.exist;
      expect(mcp!.defaultValue).to.be.true;
    });
  });

  describe("createDefaultRegistry()", () => {
    it("should return a registry pre-populated with built-in flags", () => {
      const registry = createDefaultRegistry();
      expect(registry.list().length).to.equal(builtinFlags.length);
    });
  });
});

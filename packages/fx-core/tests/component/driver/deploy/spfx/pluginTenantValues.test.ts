// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { describe, expect, it } from "vitest";
import {
  findUnresolvedTenantValue,
  resolveTenantOrigin,
  restoreTenantOrigin,
} from "../../../../../src/component/driver/deploy/spfx/utility/pluginTenantValues";

const origin = "https://tenant.example.test";

function runtimeWithTools(tools: unknown): Record<string, unknown> {
  return { spec: { mcp_tool_description: { tools } } };
}

function runtimeWithPolicy(policy: unknown): Record<string, unknown> {
  return runtimeWithTools([{ _meta: { "ui/csp": policy } }]);
}

function nested(levels: number): unknown {
  let value: unknown = "leaf";
  for (let level = 0; level < levels; level += 1) {
    value = [value];
  }
  return value;
}

const malformedRuntimes: { name: string; runtime: Record<string, unknown> }[] = [
  { name: "no spec", runtime: {} },
  { name: "a non-record spec", runtime: { spec: "{{TENANT_ORIGIN}}" } },
  { name: "no tool description", runtime: { spec: {} } },
  { name: "a non-record tool description", runtime: { spec: { mcp_tool_description: [] } } },
  { name: "a non-array tool list", runtime: runtimeWithTools({ frameDomains: [origin] }) },
  { name: "non-record tools", runtime: runtimeWithTools(["{{TENANT_ORIGIN}}", null]) },
  { name: "a tool without metadata", runtime: runtimeWithTools([{ name: "tool" }]) },
  { name: "a non-record policy", runtime: runtimeWithPolicy([origin, "{{TENANT_ORIGIN}}"]) },
  {
    name: "non-array policy lists",
    runtime: runtimeWithPolicy({ frameDomains: origin, resourceDomains: {} }),
  },
];

describe("SPFX-PREP-24: plugin tenant values", () => {
  describe("resolveTenantOrigin", () => {
    it("resolves whole-value tokens only in the qualified lists and keeps every other entry", () => {
      const runtime = runtimeWithPolicy({
        frameDomains: ["{{ TENANT_ORIGIN }}", "https://static.example.test", 7],
        resourceDomains: ["{{TENANT_ORIGIN}}"],
        connectDomains: ["{{TENANT_ORIGIN}}"],
      });

      expect(resolveTenantOrigin(runtime, origin)).toBe(true);

      expect(runtime).toEqual(
        runtimeWithPolicy({
          frameDomains: [origin, "https://static.example.test", 7],
          resourceDomains: [origin],
          connectDomains: ["{{TENANT_ORIGIN}}"],
        })
      );
    });

    it.each(["{{OTHER_VALUE}}", "https://{{TENANT_ORIGIN}}", "{{TENANT_ORIGIN}}/path"])(
      "refuses the qualified entry %s",
      (entry) => {
        expect(resolveTenantOrigin(runtimeWithPolicy({ resourceDomains: [entry] }), origin)).toBe(
          false
        );
      }
    );

    it.each(malformedRuntimes)("leaves a runtime with $name untouched", ({ runtime }) => {
      const before = structuredClone(runtime);

      expect(resolveTenantOrigin(runtime, origin)).toBe(true);

      expect(runtime).toEqual(before);
    });
  });

  describe("restoreTenantOrigin", () => {
    it("restores only the entries that equal the origin in the qualified lists", () => {
      const runtime = runtimeWithPolicy({
        frameDomains: [origin, "https://static.example.test", `${origin}/path`],
        resourceDomains: [origin],
        connectDomains: [origin],
      });

      restoreTenantOrigin(runtime, origin);

      expect(runtime).toEqual(
        runtimeWithPolicy({
          frameDomains: ["{{TENANT_ORIGIN}}", "https://static.example.test", `${origin}/path`],
          resourceDomains: ["{{TENANT_ORIGIN}}"],
          connectDomains: [origin],
        })
      );
    });

    it.each(malformedRuntimes)("leaves a runtime with $name untouched", ({ runtime }) => {
      const before = structuredClone(runtime);

      restoreTenantOrigin(runtime, origin);

      expect(runtime).toEqual(before);
    });
  });

  describe("findUnresolvedTenantValue", () => {
    it.each([
      { name: "an object member", plugin: { a: { b: "{{OTHER_VALUE}}" } }, token: "OTHER_VALUE" },
      {
        name: "an array element",
        plugin: { a: [1, ["{{ SPACED_TOKEN }}"]] },
        token: "SPACED_TOKEN",
      },
      { name: "the root", plugin: "{{ROOT_VALUE}}", token: "ROOT_VALUE" },
    ])("reports a whole-value placeholder in $name", ({ plugin, token }) => {
      expect(findUnresolvedTenantValue(plugin)).toEqual({ kind: "placeholder", token });
    });

    it("visits earlier members before later ones", () => {
      const plugin = { first: ["{{FIRST_VALUE}}"], second: "{{SECOND_VALUE}}" };

      expect(findUnresolvedTenantValue(plugin)).toEqual({
        kind: "placeholder",
        token: "FIRST_VALUE",
      });
    });

    it("preserves partial placeholders, keys, lowercase names and non-string values", () => {
      const plugin = {
        "{{KEY_NAME}}": 1,
        text: "Uses {{TENANT_ORIGIN}} inside a sentence",
        lowercase: "{{tenant_origin}}",
        empty: "{{}}",
        numbers: [1, 2.5, true, false, null],
        nested: { list: ["https://{{TENANT_ORIGIN}}"] },
      };

      expect(findUnresolvedTenantValue(plugin)).toBeUndefined();
    });

    it("accepts the deepest supported nesting and refuses anything deeper", () => {
      expect(findUnresolvedTenantValue(nested(128))).toBeUndefined();
      expect(findUnresolvedTenantValue(nested(129))).toEqual({ kind: "tooDeep" });
    });
  });
});

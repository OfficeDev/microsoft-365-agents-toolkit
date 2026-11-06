// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { describe, expect, it, vi } from "vitest";

vi.mock("../../src/component/driver/m365/acquire", () => {
  throw new Error("Metadata must not load the installation driver");
});
vi.mock("../../src/component/m365/packageService", () => {
  throw new Error("Metadata must not load the live service");
});
vi.mock("../../src/component/driver/workiq/publishPolicy", () => {
  throw new Error("Metadata must not load the policy driver");
});
vi.mock("../../src/component/workiq/policyPublication", () => {
  throw new Error("Metadata must not load publication operations");
});
vi.mock("../../src/client/workIqPolicyClient", () => {
  throw new Error("Metadata must not load the Work IQ client");
});

describe("standalone action contracts", () => {
  it("ACTION-01: exports readonly literal action names without loading live drivers", async () => {
    const { ActionNames, getActionDefinition, validateActionArgs } =
      await import("../../src/actions");
    expect(ActionNames).toEqual({
      Acquire: "teamsApp/extendToM365",
      PublishPolicy: "copilotAgent/publishPolicy",
    });
    expect(Object.isFrozen(ActionNames)).toBe(true);
    expect(getActionDefinition(ActionNames.Acquire).isOk()).toBe(true);
    expect(validateActionArgs(ActionNames.Acquire, { appPackagePath: "app.zip" }).isOk()).toBe(
      true
    );
  });

  it("ACTION-02: validates raw args without coercion, mutation or sensitive values", async () => {
    const { validateActionArgs, getActionDefinition } = await import("../../src/actions");
    const input = { appPackagePath: "app.zip", scope: "Personal" };
    const valid = validateActionArgs("teamsApp/extendToM365", input);
    expect(valid.isOk()).toBe(true);
    if (valid.isOk()) expect(valid.value).not.toBe(input);
    expect(input).toEqual({ appPackagePath: "app.zip", scope: "Personal" });

    const cases: Array<[unknown, string, string]> = [
      [{}, "/appPackagePath", "required"],
      [{ appPackagePath: 123 }, "/appPackagePath", "type"],
      [{ appPackagePath: "app.zip", scope: "secret-invalid-scope" }, "/scope", "enum"],
      [
        { appPackagePath: "app.zip", "extra/field": "secret-value" },
        "/extra~1field",
        "additionalProperties",
      ],
      [{ appPackagePath: "" }, "/appPackagePath", "minLength"],
      [{ appPackagePath: "app.zip", scope: null }, "/scope", "enum"],
    ];
    for (const [args, path, code] of cases) {
      const result = validateActionArgs("teamsApp/extendToM365", args);
      expect(result.isErr()).toBe(true);
      if (result.isErr()) {
        expect(result.error.issues).toEqual(
          expect.arrayContaining([expect.objectContaining({ path, code })])
        );
        expect(JSON.stringify(result.error)).not.toContain("secret-");
        expect(result.error.message).not.toContain("secret-");
      }
    }

    const definition = getActionDefinition("teamsApp/extendToM365");
    expect(definition.isOk()).toBe(true);
    if (definition.isOk()) {
      expect(definition.value.argsSchema.$schema).toBe("http://json-schema.org/draft-07/schema#");
      Object.assign(definition.value.argsSchema, { additionalProperties: true });
    }
    expect(
      validateActionArgs("teamsApp/extendToM365", {
        appPackagePath: "app.zip",
        extra: true,
      }).isErr()
    ).toBe(true);
    expect(getActionDefinition("constructor").isErr()).toBe(true);
    expect(getActionDefinition("teamsApp/uninstallFromM365").isErr()).toBe(true);
    const invalidName = Reflect.apply(getActionDefinition, undefined, [{ toString: "invalid" }]);
    expect(invalidName.isErr()).toBe(true);
    const unknown = Reflect.apply(validateActionArgs, undefined, ["script", {}]);
    expect(unknown.isErr()).toBe(true);
    const uncloneable = validateActionArgs("teamsApp/extendToM365", {
      appPackagePath: () => "secret-function",
    });
    expect(uncloneable.isErr()).toBe(true);
    if (uncloneable.isErr()) expect(uncloneable.error.message).not.toContain("secret-function");
  });

  it("POLICY-ACTION-01: exports policy metadata and validates args without live dependencies", async () => {
    const { ActionNames, getActionDefinition, validateActionArgs } =
      await import("../../src/actions");
    const definition = getActionDefinition(ActionNames.PublishPolicy);
    expect(definition.isOk()).toBe(true);
    expect(getActionDefinition("copilotAgent/publishPolicy").isOk()).toBe(true);
    expect(getActionDefinition("workiq/publishPolicy").isErr()).toBe(true);
    if (definition.isOk()) expect(definition.value.supportsCancellation).toBe(true);
    const valid = {
      titleId: "U_title",
      policyPath: "policy.json",
      callerId: "11111111-2222-4333-8444-555555555555",
    };
    expect(validateActionArgs(ActionNames.PublishPolicy, valid).isOk()).toBe(true);
    for (const input of [
      { titleId: "U_title", policyPath: "policy.json" },
      { ...valid, titleId: 123 },
      { ...valid, policyPath: "   " },
      { ...valid, callerId: "" },
      { ...valid, etag: '"secret-etag"' },
      { ...valid, policySnapshot: {} },
    ]) {
      const checked = validateActionArgs(ActionNames.PublishPolicy, input);
      expect(checked.isErr()).toBe(true);
      if (checked.isErr()) expect(JSON.stringify(checked.error)).not.toContain("secret-etag");
    }
  });
  it("ACTION-03: imports metadata and validates without live drivers or Tools", async () => {
    const actions = await import("../../src/actions");
    expect(
      actions.validateActionArgs("teamsApp/extendToM365", { appPackagePath: "app.zip" }).isOk()
    ).toBe(true);
    const definition = actions.getActionDefinition("teamsApp/extendToM365");
    expect(definition.isOk()).toBe(true);
    if (definition.isOk()) expect(definition.value.supportsCancellation).toBe(false);
  });

  it("CLI-ACTION-07: describes all allowlisted actions and their fields without live dependencies", async () => {
    const { ActionNames, getActionDefinition } = await import("../../src/actions");
    for (const name of Object.values(ActionNames)) {
      const definition = getActionDefinition(name);
      expect(definition.isOk()).toBe(true);
      if (definition.isErr()) continue;
      const schema = definition.value.argsSchema;
      expect(schema.description).toEqual(expect.any(String));
      expect(schema.description).not.toBe("");
      const properties = schema.properties;
      expect(properties).toBeTypeOf("object");
      if (typeof properties !== "object" || properties === null) continue;
      for (const property of Object.values(properties)) {
        expect(property).toHaveProperty("description", expect.any(String));
        expect(property.description).not.toBe("");
      }
    }
  });
});

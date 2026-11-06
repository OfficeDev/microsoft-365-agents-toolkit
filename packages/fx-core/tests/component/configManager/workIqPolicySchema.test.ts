// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "reflect-metadata";
import { Ajv } from "ajv";
import { Container } from "typedi";
import { describe, expect, it } from "vitest";
import currentSchema from "../../../resource/yaml-schema/yaml.schema.json";
import versionedSchema from "../../../resource/yaml-schema/v1.13/yaml.schema.json";
import { ActionNames, getActionDefinition } from "../../../src/actions";
import type { PublishPolicyDriver } from "../../../src/component/driver/workiq/publishPolicy";
import "../../../src/component/driver";

const inputs = { titleId: "title-id", policyPath: "./policy.json", callerId: "runtime-app-id" };
const description =
  "Publish a policy for an existing agent. Currently supports Work IQ policies only. Publication does not verify runtime enforcement.";
const fieldDescriptions = {
  titleId: {
    description: "Title ID of the existing agent. This action does not create the agent or title.",
  },
  policyPath: {
    description:
      "Path to a UTF-8 Work IQ policy JSON file inside the project directory, either absolute or relative to the project root. File contents are not expanded or rewritten.",
  },
  callerId: {
    description:
      "Nonzero App ID GUID of the agent's runtime application. Must match appId if the policy already contains it. Do not use a user object ID, tenant ID, or Toolkit client ID.",
  },
};

it("WIQ-POL-01: registers the native action in the shared driver registry", () => {
  expect(Container.has("copilotAgent/publishPolicy")).toBe(true);
  expect(Container.has("workiq/publishPolicy")).toBe(false);
});

for (const [name, schema] of [
  ["current", currentSchema],
  ["v1.13", versionedSchema],
] as const) {
  describe(`copilotAgent/publishPolicy ${name} schema`, () => {
    const ajv = new Ajv({ allowUnionTypes: true });
    ajv.addKeyword("deprecationMessage");
    const validate = ajv.compile(schema);
    const document = (withInputs: Record<string, unknown>, outputs?: Record<string, string>) => ({
      version: "v1.13",
      publish: [
        {
          uses: "copilotAgent/publishPolicy",
          with: withInputs,
          ...(outputs ? { writeToEnvironmentFile: outputs } : {}),
        },
      ],
    });

    it("WIQ-POL-01: aligns YAML, SDK and driver copy with the supported policy scope", () => {
      const action = schema.definitions.copilotAgentPublishPolicy;
      expect(action.description).toBe(description);
      expect(action.properties.uses.description).toBe(description);
      expect(action.properties.with.properties).toMatchObject(fieldDescriptions);

      const definition = getActionDefinition(ActionNames.PublishPolicy);
      expect(definition.isOk()).toBe(true);
      if (definition.isOk()) {
        expect(definition.value.argsSchema.description).toContain(description);
        expect(definition.value.argsSchema.properties).toMatchObject(fieldDescriptions);
      }

      const driver = Container.get<PublishPolicyDriver>(ActionNames.PublishPolicy);
      expect(driver.description).toBe(description);
      expect(driver.progressTitle).toBe("Publishing agent policy");
    });

    it("WIQ-POL-01: accepts explicit inputs without requiring output mappings", () => {
      expect(validate(document(inputs))).toBe(true);
    });

    it("WIQ-POL-01: rejects the previous action spelling", () => {
      const previous = document(inputs);
      previous.publish[0].uses = "workiq/publishPolicy";
      expect(validate(previous)).toBe(false);
    });

    it("WIQ-POL-01: does not expose report output mappings in YAML", () => {
      expect(validate(document(inputs, { report: "POLICY_REPORT" }))).toBe(false);
      expect(validate(document(inputs, {}))).toBe(false);
    });

    for (const key of ["titleId", "policyPath", "callerId"]) {
      it(`WIQ-POL-02: rejects a missing ${key}`, () => {
        expect(
          validate(document(Object.fromEntries(Object.entries(inputs).filter(([k]) => k !== key))))
        ).toBe(false);
      });

      it(`WIQ-POL-02: rejects an empty ${key}`, () => {
        expect(validate(document({ ...inputs, [key]: "" }))).toBe(false);
      });
    }

    for (const key of ["etag", "workIqEtag"]) {
      it(`WIQ-POL-02: rejects legacy user input ${key}`, () => {
        expect(validate(document({ ...inputs, [key]: '"stored"' }))).toBe(false);
      });
    }

    it("WIQ-POL-02: rejects unrecognized inputs and output mappings", () => {
      expect(validate(document({ ...inputs, runtimeAppId: "guessed-id" }))).toBe(false);
      expect(validate(document(inputs, { success: "PUBLISHED" }))).toBe(false);
    });
  });
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "reflect-metadata";
import { Ajv } from "ajv";
import { Container } from "typedi";
import { describe, expect, it } from "vitest";
import currentSchema from "../../../resource/yaml-schema/yaml.schema.json";
import versionedSchema from "../../../resource/yaml-schema/v1.13/yaml.schema.json";
import "../../../src/component/driver";

const inputs = { titleId: "title-id", policyPath: "./policy.json", callerId: "runtime-app-id" };

it("WIQ-POL-01: registers the native action in the shared driver registry", () => {
  expect(Container.has("workiq/publishPolicy")).toBe(true);
});

for (const [name, schema] of [
  ["current", currentSchema],
  ["v1.13", versionedSchema],
] as const) {
  describe(`workiq/publishPolicy ${name} schema`, () => {
    const ajv = new Ajv({ allowUnionTypes: true });
    ajv.addKeyword("deprecationMessage");
    const validate = ajv.compile(schema);
    const document = (withInputs: Record<string, unknown>, outputs?: Record<string, string>) => ({
      version: "v1.13",
      publish: [
        {
          uses: "workiq/publishPolicy",
          with: withInputs,
          ...(outputs ? { writeToEnvironmentFile: outputs } : {}),
        },
      ],
    });

    it("WIQ-POL-01: accepts explicit inputs without requiring output mappings", () => {
      expect(validate(document(inputs))).toBe(true);
    });

    it("WIQ-POL-01: accepts explicit replacement and an optional report mapping", () => {
      expect(
        validate(document({ ...inputs, etag: '"version-1"' }, { report: "POLICY_REPORT" }))
      ).toBe(true);
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

    it("WIQ-POL-23: accepts different MOS and Work IQ ETags", () => {
      expect(validate(document({ ...inputs, etag: '"mos"', workIqEtag: '"runtime"' }))).toBe(true);
      expect(validate(document({ ...inputs, workIqEtag: "*" }))).toBe(false);
    });

    it("WIQ-POL-02: rejects unrecognized inputs and output mappings", () => {
      expect(validate(document({ ...inputs, runtimeAppId: "guessed-id" }))).toBe(false);
      expect(validate(document(inputs, { success: "PUBLISHED" }))).toBe(false);
    });

    for (const etag of [
      'W/"weak"',
      "*",
      '"non-latin-\u0100"',
      '"line\r\nbreak"',
      '"trailing-newline"\n',
    ]) {
      it(`WIQ-POL-02: rejects a non-strong HTTP ETag ${JSON.stringify(etag)}`, () => {
        expect(validate(document({ ...inputs, etag }))).toBe(false);
      });
    }
  });
}

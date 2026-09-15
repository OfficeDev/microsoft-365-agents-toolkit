// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import Ajv from "ajv";
import fs from "fs";
import path from "path";
import { assert } from "vitest";

describe("descriptor language presentation schema", () => {
  it("CLEAN-02/07: accepts bounded presentation metadata and rejects extra behavior", () => {
    const schema = JSON.parse(
      fs.readFileSync(
        path.resolve(__dirname, "../../../../../templates/v4/schema/descriptor.schema.json"),
        "utf8"
      )
    );
    const validate = new Ajv({ strict: false }).compile(schema);
    const descriptor = {
      id: "synthetic",
      name: "Synthetic",
      languages: ["python"],
      minEngineVersion: "6.12.0",
      optionsSchema: { type: "object", properties: {} },
      replaceMap: [],
      languageOptions: [{ id: "python", label: "Python", description: "Preview" }],
    };
    assert.isTrue(validate(descriptor), JSON.stringify(validate.errors));
    assert.isFalse(
      validate({ ...descriptor, languageOptions: [{ id: "python", description: false }] })
    );
    assert.isFalse(
      validate({ ...descriptor, languageOptions: [{ id: "python", condition: "true" }] })
    );
    assert.isFalse(validate({ ...descriptor, languageOptions: [{ id: "python", keyPrefix: 3 }] }));
  });
});

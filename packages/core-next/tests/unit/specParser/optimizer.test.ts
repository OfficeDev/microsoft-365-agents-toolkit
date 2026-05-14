// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { expect } from "chai";
import { describe, it } from "mocha";
import type { OpenAPIV3 } from "openapi-types";
import { optimizeSpec } from "../../../src/specParser/optimizer";

function makeSpec(): OpenAPIV3.Document {
  return {
    openapi: "3.0.0",
    info: { title: "Test", version: "1.0.0" },
    "x-custom": "vendor extension",
    servers: [{ url: "https://api.example.com" }],
    paths: {
      "/pets": {
        get: {
          operationId: "listPets",
          tags: ["pets"],
          security: [{ apiKey1: [] }],
          responses: {
            "200": {
              description: "OK",
              content: {
                "application/json": {
                  schema: { $ref: "#/components/schemas/Pet" },
                },
              },
            },
          },
        },
      },
    },
    tags: [{ name: "pets" }, { name: "unused-tag" }],
    components: {
      schemas: {
        Pet: {
          type: "object",
          properties: { name: { type: "string" } },
        },
        UnusedSchema: {
          type: "object",
          properties: { foo: { type: "string" } },
        },
      },
      securitySchemes: {
        apiKey1: { type: "apiKey", name: "x-api-key", in: "header" },
        unusedAuth: { type: "http", scheme: "bearer" },
      },
    },
  } as unknown as OpenAPIV3.Document;
}

describe("specParser/optimizer", () => {
  it("removes vendor extensions (x-*)", () => {
    const result = optimizeSpec(makeSpec());
    expect(result).to.not.have.property("x-custom");
  });

  it("removes unused tags", () => {
    const result = optimizeSpec(makeSpec());
    expect(result.tags).to.have.length(1);
    expect(result.tags![0].name).to.equal("pets");
  });

  it("removes unused security schemes", () => {
    const result = optimizeSpec(makeSpec());
    expect(result.components?.securitySchemes).to.have.property("apiKey1");
    expect(result.components?.securitySchemes).to.not.have.property("unusedAuth");
  });

  it("removes unused component schemas", () => {
    const result = optimizeSpec(makeSpec());
    expect(result.components?.schemas).to.have.property("Pet");
    expect(result.components?.schemas).to.not.have.property("UnusedSchema");
  });

  it("preserves original spec (returns deep copy)", () => {
    const original = makeSpec();
    optimizeSpec(original);
    // Original should still have vendor extensions
    expect(original).to.have.property("x-custom");
  });

  it("respects options to skip optimization steps", () => {
    const result = optimizeSpec(makeSpec(), {
      removeUnusedTags: false,
      removeUserDefinedRootProperty: false,
    });
    expect(result).to.have.property("x-custom");
    expect(result.tags).to.have.length(2);
  });
});

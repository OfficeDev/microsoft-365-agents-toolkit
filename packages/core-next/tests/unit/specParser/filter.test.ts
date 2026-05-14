// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { expect } from "chai";
import { describe, it } from "mocha";
import { filterSpec } from "../../../src/specParser/filter";
import { ProjectType } from "../../../src/specParser/types";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function makeSpec(): any {
  return {
    openapi: "3.0.0",
    info: { title: "Test", version: "1.0.0" },
    servers: [{ url: "https://api.example.com" }],
    paths: {
      "/pets": {
        get: {
          operationId: "listPets",
          summary: "List pets",
          parameters: [{ name: "limit", in: "query", required: true, schema: { type: "integer" } }],
          responses: {
            "200": {
              description: "OK",
              content: {
                "application/json": {
                  schema: { type: "array", items: { type: "object" } },
                },
              },
            },
          },
        },
        post: {
          operationId: "createPet",
          summary: "Create a pet",
          requestBody: {
            content: {
              "application/json": {
                schema: { type: "object", properties: { name: { type: "string" } } },
              },
            },
          },
          responses: {
            "201": {
              description: "Created",
              content: { "application/json": { schema: { type: "object" } } },
            },
          },
        },
      },
      "/dogs": {
        get: {
          operationId: "listDogs",
          summary: "List dogs",
          parameters: [{ name: "breed", in: "query", schema: { type: "string" } }],
          responses: {
            "200": {
              description: "OK",
              content: {
                "application/json": { schema: { type: "array", items: { type: "object" } } },
              },
            },
          },
        },
      },
    },
    tags: [{ name: "pets" }, { name: "unused-tag" }],
  };
}

describe("specParser/filter", () => {
  const baseOpts = {
    projectType: ProjectType.SME,
    allowMethods: ["get", "post"],
    allowMissingId: true,
  };

  it("filters to only selected operations", () => {
    const spec = makeSpec();
    const filtered = filterSpec(["GET /pets"], spec, spec, baseOpts);
    expect(filtered.paths).to.have.property("/pets");
    expect(filtered.paths!["/pets"]).to.have.property("get");
    expect(filtered.paths!["/pets"]).to.not.have.property("post");
    expect(filtered.paths).to.not.have.property("/dogs");
  });

  it("includes multiple selected operations", () => {
    const spec = makeSpec();
    const filtered = filterSpec(["GET /pets", "GET /dogs"], spec, spec, baseOpts);
    expect(filtered.paths).to.have.property("/pets");
    expect(filtered.paths).to.have.property("/dogs");
  });

  it("auto-generates operationId for operations without one", () => {
    const spec = makeSpec();
    delete spec.paths["/pets"]!.get!.operationId;
    const filtered = filterSpec(["GET /pets"], spec, spec, baseOpts);
    const getOp = filtered.paths!["/pets"]!.get as any;
    expect(getOp.operationId).to.be.a("string");
    expect(getOp.operationId!.length).to.be.greaterThan(0);
  });

  it("returns empty paths for empty filter", () => {
    const spec = makeSpec();
    const filtered = filterSpec([], spec, spec, baseOpts);
    expect(Object.keys(filtered.paths!)).to.have.length(0);
  });

  it("rejects __proto__ path to prevent prototype pollution", () => {
    const spec = makeSpec();
    const filtered = filterSpec(["GET __proto__"], spec, spec, baseOpts);
    expect(Object.keys(filtered.paths!)).to.have.length(0);

    expect((filtered as any).__proto__).to.equal(Object.prototype);
  });
});

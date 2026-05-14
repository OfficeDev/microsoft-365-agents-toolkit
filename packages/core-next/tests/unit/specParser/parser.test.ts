// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { expect } from "chai";
import { describe, it } from "mocha";
import * as path from "path";
import * as fs from "fs";
import * as os from "os";
import { parseSpec, resolveEnvVars, hasCircularRefs } from "../../../src/specParser/parser";

// Minimal valid OpenAPI 3.0 spec
const VALID_SPEC_JSON = JSON.stringify({
  openapi: "3.0.0",
  info: { title: "Test", version: "1.0.0" },
  servers: [{ url: "https://api.example.com" }],
  paths: {
    "/pets": {
      get: {
        operationId: "listPets",
        summary: "List pets",
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
    },
  },
});

// Swagger 2.0 spec
const SWAGGER2_SPEC_JSON = JSON.stringify({
  swagger: "2.0",
  info: { title: "Test", version: "1.0.0" },
  host: "api.example.com",
  basePath: "/",
  schemes: ["https"],
  paths: {
    "/pets": {
      get: {
        operationId: "listPets",
        summary: "List pets",
        produces: ["application/json"],
        responses: {
          "200": {
            description: "OK",
            schema: { type: "array", items: { type: "object" } },
          },
        },
      },
    },
  },
});

function writeTmpSpec(content: string, ext = ".json"): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "spec-parser-test-"));
  const filePath = path.join(dir, `spec${ext}`);
  fs.writeFileSync(filePath, content, "utf-8");
  return filePath;
}

describe("specParser/parser", () => {
  describe("parseSpec", () => {
    it("should parse a valid OpenAPI 3.0 spec from file", async () => {
      const filePath = writeTmpSpec(VALID_SPEC_JSON);
      const result = await parseSpec(filePath);
      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.spec.unresolved.openapi).to.equal("3.0.0");
        expect(result.value.spec.resolved.openapi).to.equal("3.0.0");
        expect(result.value.spec.isConverted).to.be.false;
        expect(result.value.spec.specHash).to.be.a("string").with.length(64);
        expect(result.value.warnings).to.have.length(0);
      }
    });

    it("should convert Swagger 2.0 to OpenAPI 3.0 when allowSwagger is true", async () => {
      const filePath = writeTmpSpec(SWAGGER2_SPEC_JSON);
      const result = await parseSpec(filePath, { allowSwagger: true });
      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.spec.isConverted).to.be.true;
        expect(result.value.spec.resolved.openapi).to.match(/^3\.0/);
        expect(result.value.warnings).to.have.length.greaterThan(0);
        expect(result.value.warnings[0].type).to.equal("convert-swagger-to-openapi");
      }
    });

    it("should return error for Swagger 2.0 when allowSwagger is false", async () => {
      const filePath = writeTmpSpec(SWAGGER2_SPEC_JSON);
      const result = await parseSpec(filePath, { allowSwagger: false });
      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.code).to.equal("SwaggerNotSupported");
      }
    });

    it("should return error for invalid spec", async () => {
      const filePath = writeTmpSpec("not valid json or yaml");
      const result = await parseSpec(filePath);
      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.code).to.equal("ParseSpecFailed");
      }
    });

    it("should return error for non-existent file", async () => {
      const result = await parseSpec("/nonexistent/path/spec.json");
      expect(result.isErr()).to.be.true;
    });
  });

  describe("resolveEnvVars", () => {
    it("should replace env var placeholders", () => {
      process.env.TEST_SPEC_HOST = "api.example.com";
      const spec = {
        openapi: "3.0.0",
        info: { title: "Test", version: "1.0.0" },
        servers: [{ url: "https://${{TEST_SPEC_HOST}}" }],
        paths: {},
      } as any;
      const resolved = resolveEnvVars(spec);
      expect(resolved.servers![0].url).to.equal("https://api.example.com");
      delete process.env.TEST_SPEC_HOST;
    });

    it("should leave unresolvable placeholders intact", () => {
      const spec = {
        openapi: "3.0.0",
        info: { title: "Test", version: "1.0.0" },
        servers: [{ url: "https://${{NONEXISTENT_VAR}}" }],
        paths: {},
      } as any;
      const resolved = resolveEnvVars(spec);
      expect(resolved.servers![0].url).to.include("${{NONEXISTENT_VAR}}");
    });
  });

  describe("hasCircularRefs", () => {
    it("should return false for non-circular spec", async () => {
      const filePath = writeTmpSpec(VALID_SPEC_JSON);
      const result = await hasCircularRefs(filePath);
      expect(result).to.be.false;
    });
  });
});

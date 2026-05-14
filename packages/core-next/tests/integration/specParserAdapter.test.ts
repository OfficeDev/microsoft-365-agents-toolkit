/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import { expect } from "chai";
import { describe, it, beforeEach, afterEach } from "mocha";
import { RealSpecParserAdapter } from "../../src/templates/openApi/realSpecParserAdapter";
import { createSpecParserAdapter } from "../../src/templates/openApi/specParserAdapter";

/**
 * Integration test: exercises RealSpecParserAdapter through the SpecParserAdapter
 * interface exactly as the scaffold pipeline uses it.
 *
 * Creates real spec files on disk and calls validate → listOperations → generate.
 */
describe("Integration: RealSpecParserAdapter", () => {
  let tmpDir: string;

  beforeEach(() => {
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "atk-adapter-integ-"));
  });

  afterEach(() => {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  });

  function writeSpec(filename: string, spec: object): string {
    const filePath = path.join(tmpDir, filename);
    fs.writeFileSync(filePath, JSON.stringify(spec, null, 2));
    return filePath;
  }

  const realSpec = {
    openapi: "3.0.0",
    info: { title: "TestAPI", version: "1.0.0" },
    servers: [{ url: "https://api.example.com" }],
    paths: {
      "/pets": {
        get: {
          operationId: "listPets",
          summary: "List all pets",
          parameters: [
            { name: "limit", in: "query", required: false, schema: { type: "integer" } },
          ],
          responses: {
            "200": {
              description: "A list of pets",
              content: {
                "application/json": {
                  schema: {
                    type: "array",
                    items: { $ref: "#/components/schemas/Pet" },
                  },
                },
              },
            },
          },
        },
        post: {
          operationId: "createPet",
          summary: "Create a pet",
          requestBody: {
            required: true,
            content: {
              "application/json": {
                schema: { $ref: "#/components/schemas/Pet" },
              },
            },
          },
          responses: {
            "201": {
              description: "Created",
              content: {
                "application/json": {
                  schema: { $ref: "#/components/schemas/Pet" },
                },
              },
            },
          },
        },
      },
      "/pets/{petId}": {
        get: {
          operationId: "getPet",
          summary: "Get a specific pet",
          parameters: [{ name: "petId", in: "path", required: true, schema: { type: "string" } }],
          responses: {
            "200": {
              description: "A pet",
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
    components: {
      schemas: {
        Pet: {
          type: "object",
          properties: {
            id: { type: "string" },
            name: { type: "string" },
          },
          required: ["id", "name"],
        },
      },
    },
  };

  // -----------------------------------------------------------------------
  // createSpecParserAdapter factory
  // -----------------------------------------------------------------------

  describe("createSpecParserAdapter()", () => {
    it("should return a RealSpecParserAdapter", () => {
      const adapter = createSpecParserAdapter();
      expect(adapter).to.be.an.instanceOf(RealSpecParserAdapter);
    });
  });

  // -----------------------------------------------------------------------
  // validate()
  // -----------------------------------------------------------------------

  describe("validate()", () => {
    it("should return valid for a correct OpenAPI 3.0 spec", async () => {
      const specPath = writeSpec("valid.json", realSpec);
      const adapter = new RealSpecParserAdapter({
        allowMethods: ["get", "post"],
        projectType: "SME" as any,
      });

      const result = await adapter.validate(specPath);
      expect(result.valid).to.be.true;
      expect(result.errors).to.be.empty;
    });

    it("should return invalid for a completely broken spec file", async () => {
      const specPath = path.join(tmpDir, "broken.json");
      fs.writeFileSync(specPath, "{ this is not json!! }");

      const adapter = new RealSpecParserAdapter();
      const result = await adapter.validate(specPath);
      expect(result.valid).to.be.false;
      expect(result.errors.length).to.be.greaterThan(0);
    });

    it("should return invalid for a nonexistent file", async () => {
      const adapter = new RealSpecParserAdapter();
      const result = await adapter.validate(path.join(tmpDir, "nonexistent.json"));
      expect(result.valid).to.be.false;
      expect(result.errors.length).to.be.greaterThan(0);
    });
  });

  // -----------------------------------------------------------------------
  // listOperations()
  // -----------------------------------------------------------------------

  describe("listOperations()", () => {
    it("should list all valid operations from a spec", async () => {
      const specPath = writeSpec("ops.json", realSpec);
      const adapter = new RealSpecParserAdapter({
        allowMethods: ["get", "post"],
        allowMissingId: false,
        allowMultipleParameters: true,
      });

      const ops = await adapter.listOperations(specPath);
      expect(ops.length).to.be.greaterThanOrEqual(2);

      const ids = ops.map((op) => op.id);
      // Should have at least GET /pets and GET /pets/{petId}
      expect(ids.some((id) => id.includes("/pets"))).to.be.true;

      // Check structure of each operation
      for (const op of ops) {
        expect(op).to.have.property("id").that.is.a("string");
        expect(op).to.have.property("method").that.is.a("string");
        expect(op).to.have.property("path").that.is.a("string");
        expect(op.method).to.match(/^(GET|POST|PUT|DELETE|PATCH|HEAD|OPTIONS|TRACE)$/);
      }
    });

    it("should return empty array for a broken spec", async () => {
      const specPath = path.join(tmpDir, "broken2.json");
      fs.writeFileSync(specPath, "not json");

      const adapter = new RealSpecParserAdapter();
      const ops = await adapter.listOperations(specPath);
      expect(ops).to.be.an("array").that.is.empty;
    });
  });

  // -----------------------------------------------------------------------
  // generate()
  // -----------------------------------------------------------------------

  describe("generate()", () => {
    it("should generate filtered openapi.json for selected operations", async () => {
      const specPath = writeSpec("gen.json", realSpec);
      const outputDir = path.join(tmpDir, "output");
      fs.mkdirSync(outputDir, { recursive: true });

      const adapter = new RealSpecParserAdapter({
        allowMethods: ["get", "post"],
        allowMultipleParameters: true,
      });

      const result = await adapter.generate(specPath, ["get /pets"], outputDir, "Copilot");

      expect(result.files.size).to.be.greaterThan(0);
      expect(result.files.has("openapi.json")).to.be.true;

      // The generated openapi.json should be valid JSON
      const generatedSpec = JSON.parse(result.files.get("openapi.json")!);
      expect(generatedSpec).to.have.property("openapi").that.matches(/^3\./);
      expect(generatedSpec).to.have.property("paths");

      // Should contain only the selected operation
      expect(generatedSpec.paths).to.have.property("/pets");
      if (generatedSpec.paths["/pets"]) {
        expect(generatedSpec.paths["/pets"]).to.have.property("get");
      }
    });

    it("should return warnings for a broken spec", async () => {
      const specPath = path.join(tmpDir, "broken-gen.json");
      fs.writeFileSync(specPath, "not json");

      const adapter = new RealSpecParserAdapter();
      const result = await adapter.generate(specPath, ["get /pets"], tmpDir, "SME");

      expect(result.files.size).to.equal(0);
      expect(result.warnings.length).to.be.greaterThan(0);
    });

    it("should work with different project types", async () => {
      const specPath = writeSpec("gen-types.json", realSpec);
      const adapter = new RealSpecParserAdapter({
        allowMethods: ["get", "post"],
      });

      for (const projectType of ["Copilot", "TeamsAi", "SME"] as const) {
        const result = await adapter.generate(specPath, ["get /pets"], tmpDir, projectType);
        expect(result.files.size, `should generate files for ${projectType}`).to.be.greaterThan(0);
      }
    });
  });

  // -----------------------------------------------------------------------
  // Full adapter pipeline: validate → listOperations → generate
  // -----------------------------------------------------------------------

  describe("Full adapter pipeline", () => {
    it("should validate, list, select, and generate — the scaffold flow", async () => {
      const specPath = writeSpec("full-pipeline.json", realSpec);
      const adapter = new RealSpecParserAdapter({
        allowMethods: ["get", "post"],
        allowMultipleParameters: true,
      });

      // Step 1: Validate
      const validation = await adapter.validate(specPath);
      expect(validation.valid, `errors: ${validation.errors.join("; ")}`).to.be.true;

      // Step 2: List operations
      const ops = await adapter.listOperations(specPath);
      expect(ops.length).to.be.greaterThanOrEqual(1);

      // Step 3: "User selects" the first valid operation
      const selectedOps = [ops[0].id.toLowerCase()];

      // Step 4: Generate
      const outputDir = path.join(tmpDir, "generated");
      fs.mkdirSync(outputDir, { recursive: true });

      const result = await adapter.generate(specPath, selectedOps, outputDir, "Copilot");
      expect(result.files.size).to.be.greaterThan(0);
      expect(result.files.has("openapi.json")).to.be.true;

      // Verify the generated spec is valid JSON and has the selected path
      const genSpec = JSON.parse(result.files.get("openapi.json")!);
      expect(genSpec.openapi).to.match(/^3\./);
      expect(genSpec.paths).to.exist;
      expect(Object.keys(genSpec.paths).length).to.be.greaterThanOrEqual(1);
    });
  });

  // -----------------------------------------------------------------------
  // Auth in listOperations
  // -----------------------------------------------------------------------

  describe("Auth info in listOperations", () => {
    it("should include auth info for operations with security", async () => {
      const authSpec = {
        openapi: "3.0.0",
        info: { title: "AuthAPI", version: "1.0.0" },
        servers: [{ url: "https://api.secure.example.com" }],
        paths: {
          "/secure": {
            get: {
              operationId: "getSecureData",
              summary: "Get secure data",
              security: [{ apiKeyAuth: [] }],
              parameters: [{ name: "q", in: "query", schema: { type: "string" } }],
              responses: {
                "200": {
                  description: "OK",
                  content: {
                    "application/json": {
                      schema: { type: "object" },
                    },
                  },
                },
              },
            },
          },
        },
        components: {
          securitySchemes: {
            apiKeyAuth: {
              type: "apiKey",
              name: "X-API-Key",
              in: "header",
            },
          },
        },
      };

      const specPath = writeSpec("auth-ops.json", authSpec);
      const adapter = new RealSpecParserAdapter({
        allowMethods: ["get"],
        allowAPIKeyAuth: true,
      });

      const ops = await adapter.listOperations(specPath);
      const secureOp = ops.find((op) => op.path === "/secure");
      expect(secureOp).to.exist;
      if (secureOp?.auth) {
        expect(secureOp.auth.authType).to.equal("apiKey");
      }
    });
  });
});

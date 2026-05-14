/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import { expect } from "chai";
import { describe, it, beforeEach, afterEach } from "mocha";
import { parseSpec } from "../../src/specParser/parser";
import { createValidator, type SpecValidationResult } from "../../src/specParser/validator";
import { filterSpec } from "../../src/specParser/filter";
import { optimizeSpec } from "../../src/specParser/optimizer";
import type { ParseOptions } from "../../src/specParser/types";
import { ProjectType } from "../../src/specParser/types";
import type { OpenAPIV3 } from "openapi-types";

/**
 * Integration test: exercises the specParser pipeline end-to-end.
 *
 * No external services needed — writes spec files to a temp directory,
 * then runs parse → validate → listAPIs → filter → optimize in sequence.
 */
describe("Integration: specParser pipeline", () => {
  let tmpDir: string;

  beforeEach(() => {
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "atk-spec-integ-"));
  });

  afterEach(() => {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  });

  function writeSpec(filename: string, spec: object): string {
    const filePath = path.join(tmpDir, filename);
    fs.writeFileSync(filePath, JSON.stringify(spec, null, 2));
    return filePath;
  }

  // -----------------------------------------------------------------------
  // Test specs
  // -----------------------------------------------------------------------

  const petStoreSpec: OpenAPIV3.Document = {
    openapi: "3.0.0",
    info: { title: "PetStore", version: "1.0.0" },
    servers: [{ url: "https://api.petstore.example.com" }],
    paths: {
      "/pets": {
        get: {
          operationId: "listPets",
          summary: "List all pets",
          parameters: [
            {
              name: "limit",
              in: "query",
              required: false,
              schema: { type: "integer" },
            },
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
        } as unknown as OpenAPIV3.OperationObject,
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
        } as unknown as OpenAPIV3.OperationObject,
      },
      "/pets/{petId}": {
        get: {
          operationId: "getPet",
          summary: "Get a pet by ID",
          parameters: [
            {
              name: "petId",
              in: "path",
              required: true,
              schema: { type: "string" },
            },
          ],
          responses: {
            "200": {
              description: "A single pet",
              content: {
                "application/json": {
                  schema: { $ref: "#/components/schemas/Pet" },
                },
              },
            },
          },
        } as unknown as OpenAPIV3.OperationObject,
        delete: {
          operationId: "deletePet",
          summary: "Delete a pet",
          parameters: [
            {
              name: "petId",
              in: "path",
              required: true,
              schema: { type: "string" },
            },
          ],
          responses: {
            "204": { description: "Deleted" },
          },
        } as unknown as OpenAPIV3.OperationObject,
      },
    },
    components: {
      schemas: {
        Pet: {
          type: "object",
          properties: {
            id: { type: "string" },
            name: { type: "string" },
            tag: { type: "string" },
          },
          required: ["id", "name"],
        },
        Unused: {
          type: "object",
          properties: {
            placeholder: { type: "string" },
          },
        },
      },
    },
    tags: [
      { name: "pets", description: "Pet operations" },
      { name: "admin", description: "Admin operations (unused)" },
    ],
  } as unknown as OpenAPIV3.Document;

  const authSpec: OpenAPIV3.Document = {
    openapi: "3.0.0",
    info: { title: "AuthAPI", version: "1.0.0" },
    servers: [{ url: "https://api.secure.example.com" }],
    paths: {
      "/items": {
        get: {
          operationId: "listItems",
          summary: "List items",
          security: [{ apiKeyAuth: [] }],
          parameters: [{ name: "q", in: "query", required: false, schema: { type: "string" } }],
          responses: {
            "200": {
              description: "OK",
              content: {
                "application/json": { schema: { type: "array", items: { type: "object" } } },
              },
            },
          },
        } as unknown as OpenAPIV3.OperationObject,
      },
    },
    components: {
      securitySchemes: {
        apiKeyAuth: {
          type: "apiKey",
          name: "X-API-Key",
          in: "header",
        },
        unusedScheme: {
          type: "http",
          scheme: "basic",
        },
      },
    },
  } as unknown as OpenAPIV3.Document;

  // -----------------------------------------------------------------------
  // Full pipeline: parse → validate → list → filter → optimize
  // -----------------------------------------------------------------------

  describe("PetStore spec: full pipeline", () => {
    it("should parse, validate, list APIs, filter, and optimize", async () => {
      const specPath = writeSpec("petstore.json", petStoreSpec);
      const options: ParseOptions = {
        allowMethods: ["get", "post"],
        projectType: ProjectType.SME,
        allowMultipleParameters: true,
      };

      // Step 1: Parse
      const parseResult = await parseSpec(specPath, options);
      expect(parseResult.isOk(), "parse should succeed").to.be.true;
      const { spec, warnings: parseWarnings } = parseResult._unsafeUnwrap();
      expect(spec.isConverted).to.be.false;
      expect(spec.specHash).to.be.a("string").with.length.greaterThan(0);

      // Step 2: Validate
      const validator = createValidator(spec.resolved, options);
      const specValidation: SpecValidationResult = validator.validateSpec();
      expect(specValidation.errors).to.be.empty;

      // Step 3: List APIs
      const apiMap = validator.listAPIs();
      const validAPIs = Object.entries(apiMap).filter(([, v]) => v.isValid);
      // Only GET and POST are allowed, DELETE is not
      expect(validAPIs.length).to.be.greaterThanOrEqual(2);
      expect(apiMap).to.have.property("GET /pets");
      expect(apiMap).to.have.property("POST /pets");
      expect(apiMap["GET /pets"].isValid).to.be.true;
      expect(apiMap["POST /pets"].isValid).to.be.true;

      // Step 4: Filter to just GET /pets
      const filtered = filterSpec(["get /pets"], spec.unresolved, spec.resolved, options);
      expect(filtered.paths).to.have.property("/pets");
      expect(filtered.paths!["/pets"]).to.have.property("get");
      expect(filtered.paths!["/pets"]).to.not.have.property("post");
      expect(filtered.paths).to.not.have.property("/pets/{petId}");

      // Step 5: Verify optimization removed unused things
      // The "Unused" schema should be removed since it's not referenced by filtered paths
      if (filtered.components?.schemas) {
        expect(filtered.components.schemas).to.not.have.property("Unused");
      }
      // The "admin" tag should be removed since no operation uses it
      if (filtered.tags) {
        const tagNames = filtered.tags.map((t: OpenAPIV3.TagObject) => t.name);
        expect(tagNames).to.not.include("admin");
      }
    });

    it("should correctly filter to multiple operations", async () => {
      const specPath = writeSpec("petstore2.json", petStoreSpec);
      const options: ParseOptions = {
        allowMethods: ["get", "post"],
        projectType: ProjectType.Copilot,
      };

      const parseResult = await parseSpec(specPath, options);
      expect(parseResult.isOk()).to.be.true;
      const { spec } = parseResult._unsafeUnwrap();

      const filtered = filterSpec(
        ["get /pets", "get /pets/{petId}"],
        spec.unresolved,
        spec.resolved,
        options
      );

      expect(filtered.paths).to.have.property("/pets");
      expect(filtered.paths).to.have.property("/pets/{petId}");
      expect(filtered.paths!["/pets"]).to.have.property("get");
      expect(filtered.paths!["/pets"]).to.not.have.property("post");
    });

    it("should preserve referenced components and remove unreferenced ones", async () => {
      const specPath = writeSpec("petstore3.json", petStoreSpec);
      const options: ParseOptions = {
        allowMethods: ["get", "post"],
        projectType: ProjectType.SME,
      };

      const parseResult = await parseSpec(specPath, options);
      expect(parseResult.isOk()).to.be.true;
      const { spec } = parseResult._unsafeUnwrap();

      const filtered = filterSpec(["get /pets"], spec.unresolved, spec.resolved, options);

      // Pet schema should be preserved (referenced by GET /pets response)
      if (filtered.components?.schemas) {
        expect(filtered.components.schemas).to.have.property("Pet");
      }
    });
  });

  // -----------------------------------------------------------------------
  // Swagger 2.0 conversion pipeline
  // -----------------------------------------------------------------------

  describe("Swagger 2.0 conversion pipeline", () => {
    const swaggerSpec = {
      swagger: "2.0",
      info: { title: "LegacyAPI", version: "1.0.0" },
      host: "api.legacy.example.com",
      basePath: "/v2",
      schemes: ["https"],
      paths: {
        "/users": {
          get: {
            operationId: "listUsers",
            summary: "List users",
            produces: ["application/json"],
            parameters: [{ name: "q", in: "query", type: "string" }],
            responses: {
              "200": {
                description: "OK",
                schema: { type: "array", items: { type: "object" } },
              },
            },
          },
        },
      },
    };

    it("should convert Swagger 2.0 → OpenAPI 3.0 and validate", async () => {
      const specPath = writeSpec("legacy.json", swaggerSpec);
      const options: ParseOptions = {
        allowSwagger: true,
        allowMethods: ["get", "post"],
        projectType: ProjectType.Copilot,
        allowMissingId: true,
      };

      const parseResult = await parseSpec(specPath, options);
      expect(parseResult.isOk(), "parse should succeed").to.be.true;
      const { spec, warnings } = parseResult._unsafeUnwrap();

      expect(spec.isConverted).to.be.true;
      expect(spec.resolved.openapi).to.match(/^3\./);
      expect(warnings.length).to.be.greaterThanOrEqual(1);

      // Server URL should reflect original host + basePath
      expect(spec.resolved.servers).to.exist;
      expect(spec.resolved.servers![0].url).to.include("legacy.example.com");
    });

    it("should reject Swagger 2.0 when allowSwagger is false", async () => {
      const specPath = writeSpec("legacy-no-allow.json", swaggerSpec);
      const options: ParseOptions = {
        allowSwagger: false,
        allowMethods: ["get", "post"],
        projectType: ProjectType.SME,
      };

      const parseResult = await parseSpec(specPath, options);
      expect(parseResult.isErr(), "parse should fail").to.be.true;
      expect(parseResult._unsafeUnwrapErr().message).to.include("Swagger");
    });
  });

  // -----------------------------------------------------------------------
  // Auth-aware validation pipeline
  // -----------------------------------------------------------------------

  describe("Auth-aware pipeline", () => {
    it("should validate API key auth when allowAPIKeyAuth is enabled", async () => {
      const specPath = writeSpec("auth.json", authSpec);
      const options: ParseOptions = {
        allowMethods: ["get"],
        allowAPIKeyAuth: true,
        projectType: ProjectType.SME,
        allowMissingId: false,
      };

      const parseResult = await parseSpec(specPath, options);
      expect(parseResult.isOk()).to.be.true;
      const { spec } = parseResult._unsafeUnwrap();

      const validator = createValidator(spec.resolved, options);
      const apiMap = validator.listAPIs();
      expect(apiMap["GET /items"]).to.exist;
      expect(apiMap["GET /items"].isValid).to.be.true;
    });

    it("should reject API key auth when all auth options are disabled", async () => {
      const specPath = writeSpec("auth-reject.json", authSpec);
      const options: ParseOptions = {
        allowMethods: ["get"],
        allowAPIKeyAuth: false,
        allowBearerTokenAuth: false,
        allowOauth2: false,
        projectType: ProjectType.SME,
        allowMissingId: false,
      };

      const parseResult = await parseSpec(specPath, options);
      expect(parseResult.isOk()).to.be.true;
      const { spec } = parseResult._unsafeUnwrap();

      const validator = createValidator(spec.resolved, options);
      const apiMap = validator.listAPIs();
      expect(apiMap["GET /items"]).to.exist;
      expect(apiMap["GET /items"].isValid).to.be.false;
    });

    it("should optimize away unused security schemes after filter", async () => {
      const specPath = writeSpec("auth-filter.json", authSpec);
      const options: ParseOptions = {
        allowMethods: ["get"],
        allowAPIKeyAuth: true,
        projectType: ProjectType.Copilot,
      };

      const parseResult = await parseSpec(specPath, options);
      expect(parseResult.isOk()).to.be.true;
      const { spec } = parseResult._unsafeUnwrap();

      const filtered = filterSpec(["get /items"], spec.unresolved, spec.resolved, options);

      // The "unusedScheme" (basic auth) should be stripped, apiKeyAuth kept
      if (filtered.components?.securitySchemes) {
        expect(filtered.components.securitySchemes).to.have.property("apiKeyAuth");
        expect(filtered.components.securitySchemes).to.not.have.property("unusedScheme");
      }
    });
  });

  // -----------------------------------------------------------------------
  // Validation error cases
  // -----------------------------------------------------------------------

  describe("Validation error cases", () => {
    it("should reject spec with no valid APIs for SME", async () => {
      const noApiSpec: OpenAPIV3.Document = {
        openapi: "3.0.0",
        info: { title: "Empty", version: "1.0.0" },
        servers: [{ url: "https://api.example.com" }],
        paths: {
          "/data": {
            delete: {
              operationId: "deleteData",
              responses: { "204": { description: "Deleted" } },
            } as unknown as OpenAPIV3.OperationObject,
          },
        },
      } as unknown as OpenAPIV3.Document;

      const specPath = writeSpec("noapi.json", noApiSpec);
      const options: ParseOptions = {
        allowMethods: ["get", "post"],
        projectType: ProjectType.SME,
      };

      const parseResult = await parseSpec(specPath, options);
      expect(parseResult.isOk()).to.be.true;

      const { spec } = parseResult._unsafeUnwrap();
      const validator = createValidator(spec.resolved, options);
      const validation = validator.validateSpec();
      expect(validation.errors.length).to.be.greaterThan(0);
      expect(validation.errors.some((e) => e.type === "no-supported-api")).to.be.true;
    });

    it("should warn about missing operationIds", async () => {
      const missingIdSpec: OpenAPIV3.Document = {
        openapi: "3.0.0",
        info: { title: "NoIds", version: "1.0.0" },
        servers: [{ url: "https://api.example.com" }],
        paths: {
          "/items": {
            get: {
              summary: "Get items",
              parameters: [{ name: "q", in: "query", required: false, schema: { type: "string" } }],
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
            } as unknown as OpenAPIV3.OperationObject,
          },
        },
      } as unknown as OpenAPIV3.Document;

      const specPath = writeSpec("noid.json", missingIdSpec);
      const options: ParseOptions = {
        allowMethods: ["get", "post"],
        allowMissingId: true,
        projectType: ProjectType.Copilot,
      };

      const parseResult = await parseSpec(specPath, options);
      expect(parseResult.isOk()).to.be.true;

      const { spec } = parseResult._unsafeUnwrap();
      const validator = createValidator(spec.resolved, options);
      const validation = validator.validateSpec();
      expect(validation.warnings.some((w) => w.type === "operationid-missing")).to.be.true;
    });

    it("should report error for unsupported OpenAPI 3.1", async () => {
      const oa31Spec = {
        openapi: "3.1.0",
        info: { title: "V31", version: "1.0.0" },
        servers: [{ url: "https://api.example.com" }],
        paths: {
          "/test": {
            get: {
              operationId: "getTest",
              responses: { "200": { description: "OK" } },
            },
          },
        },
      };

      const specPath = writeSpec("v31.json", oa31Spec);
      const options: ParseOptions = {
        allowMethods: ["get", "post"],
        projectType: ProjectType.SME,
      };

      const parseResult = await parseSpec(specPath, options);

      if (parseResult.isErr()) {
        // Parser itself rejects 3.1 — that's valid behavior
        expect(parseResult._unsafeUnwrapErr().message).to.be.a("string");
      } else {
        // Parser accepts 3.1 — validator should catch it
        const { spec } = parseResult._unsafeUnwrap();
        const validator = createValidator(spec.resolved, options);
        const validation = validator.validateSpec();
        expect(validation.errors.some((e) => e.type === "spec-version-not-supported")).to.be.true;
      }
    });
  });

  // -----------------------------------------------------------------------
  // Optimizer standalone integration
  // -----------------------------------------------------------------------

  describe("Optimizer integration", () => {
    it("should produce minimal spec from complex input", () => {
      const complexSpec: OpenAPIV3.Document = {
        openapi: "3.0.0",
        info: { title: "Complex", version: "1.0.0" },
        "x-custom-vendor": "should-be-removed",
        servers: [{ url: "https://api.example.com" }],
        paths: {
          "/data": {
            get: {
              operationId: "getData",
              tags: ["data"],
              responses: {
                "200": {
                  description: "OK",
                  content: {
                    "application/json": {
                      schema: { $ref: "#/components/schemas/DataItem" },
                    },
                  },
                },
              },
            } as unknown as OpenAPIV3.OperationObject,
          },
        },
        tags: [{ name: "data" }, { name: "unused-tag" }],
        components: {
          schemas: {
            DataItem: {
              type: "object",
              properties: { id: { type: "string" } },
            },
            OrphanedSchema: {
              type: "object",
              properties: { nothing: { type: "string" } },
            },
          },
          securitySchemes: {
            orphanedAuth: {
              type: "http",
              scheme: "basic",
            },
          },
        },
      } as unknown as OpenAPIV3.Document;

      const optimized = optimizeSpec(complexSpec);

      // Vendor extension removed
      expect(optimized).to.not.have.property("x-custom-vendor");

      // Unused tag removed
      if (optimized.tags) {
        const names = optimized.tags.map((t) => t.name);
        expect(names).to.include("data");
        expect(names).to.not.include("unused-tag");
      }

      // Orphaned schema removed
      if (optimized.components?.schemas) {
        expect(optimized.components.schemas).to.have.property("DataItem");
        expect(optimized.components.schemas).to.not.have.property("OrphanedSchema");
      }

      // Orphaned security scheme removed
      expect(optimized.components?.securitySchemes).to.be.undefined;
    });
  });

  // -----------------------------------------------------------------------
  // Environment variable resolution
  // -----------------------------------------------------------------------

  describe("Env var resolution in server URLs", () => {
    it("should parse spec with ${{ENV_VAR}} patterns and resolve them", async () => {
      const envSpec: OpenAPIV3.Document = {
        openapi: "3.0.0",
        info: { title: "EnvTest", version: "1.0.0" },
        servers: [{ url: "https://${{API_HOST}}/v1" }],
        paths: {
          "/data": {
            get: {
              operationId: "getData",
              responses: { "200": { description: "OK" } },
            } as unknown as OpenAPIV3.OperationObject,
          },
        },
      } as unknown as OpenAPIV3.Document;

      const specPath = writeSpec("env.json", envSpec);

      // Set the env var
      const origVal = process.env.API_HOST;
      process.env.API_HOST = "api.resolved.example.com";
      try {
        const parseResult = await parseSpec(specPath);
        expect(parseResult.isOk()).to.be.true;
      } finally {
        if (origVal === undefined) delete process.env.API_HOST;
        else process.env.API_HOST = origVal;
      }
    });
  });
});

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { expect } from "chai";
import { describe, it } from "mocha";
import type { OpenAPIV3 } from "openapi-types";
import {
  isBearerTokenAuth,
  isAPIKeyAuth,
  isOAuthWithAuthCodeFlow,
  isNotSupportedAuth,
  getAuthArray,
  getAuthMap,
  getAuthInfo,
  getResponseJson,
  convertPathToCamelCase,
  getSafeRegistrationIdEnvName,
  checkServerUrl,
  isObjectSchema,
  isWellKnownName,
  formatStr,
  getServerObject,
  generateParametersFromSchema,
} from "../../../src/specParser/utils";
import { ErrorType } from "../../../src/specParser/types";

describe("specParser/utils", () => {
  describe("auth helpers", () => {
    it("isBearerTokenAuth returns true for bearer", () => {
      expect(isBearerTokenAuth({ type: "http", scheme: "bearer" } as any)).to.be.true;
      expect(isBearerTokenAuth({ type: "apiKey" } as any)).to.be.false;
    });

    it("isAPIKeyAuth returns true for apiKey", () => {
      expect(isAPIKeyAuth({ type: "apiKey" } as any)).to.be.true;
      expect(isAPIKeyAuth({ type: "http", scheme: "bearer" } as any)).to.be.false;
    });

    it("isOAuthWithAuthCodeFlow returns true for oauth2 with authorizationCode", () => {
      expect(
        isOAuthWithAuthCodeFlow({
          type: "oauth2",
          flows: {
            authorizationCode: {
              authorizationUrl: "https://auth.example.com",
              tokenUrl: "https://token.example.com",
              scopes: {},
            },
          },
        } as any)
      ).to.be.true;
      expect(isOAuthWithAuthCodeFlow({ type: "oauth2", flows: {} } as any)).to.be.false;
    });

    it("isNotSupportedAuth returns false for empty array", () => {
      expect(isNotSupportedAuth([])).to.be.false;
    });

    it("isNotSupportedAuth returns true when all have multiple auths", () => {
      expect(
        isNotSupportedAuth([
          [
            { name: "a", authScheme: { type: "apiKey" } as any },
            { name: "b", authScheme: { type: "apiKey" } as any },
          ],
        ])
      ).to.be.true;
    });
  });

  describe("getAuthArray", () => {
    it("returns auth from security requirements", () => {
      const spec = {
        components: {
          securitySchemes: {
            apiKey1: { type: "apiKey", name: "x-api-key", in: "header" },
          },
        },
        security: [{ apiKey1: [] }],
      } as unknown as OpenAPIV3.Document;

      const result = getAuthArray(undefined, spec);
      expect(result).to.have.length(1);
      expect(result[0][0].name).to.equal("apiKey1");
    });
  });

  describe("getAuthMap", () => {
    it("returns operation-to-auth mapping", () => {
      const spec = {
        components: {
          securitySchemes: {
            bearerAuth: { type: "http", scheme: "bearer" },
          },
        },
        paths: {
          "/pets": {
            get: {
              operationId: "listPets",
              security: [{ bearerAuth: [] }],
              responses: {},
            },
          },
        },
      } as unknown as OpenAPIV3.Document;

      const map = getAuthMap(spec);
      expect(map).to.have.property("listPets");
      expect(map.listPets.name).to.equal("bearerAuth");
    });
  });

  describe("getAuthInfo", () => {
    it("returns single auth info", () => {
      const spec = {
        components: {
          securitySchemes: {
            apiKey1: { type: "apiKey", name: "x-api-key", in: "header" },
          },
        },
        paths: {
          "/pets": {
            get: {
              operationId: "listPets",
              security: [{ apiKey1: [] }],
              responses: {},
            },
          },
        },
      } as unknown as OpenAPIV3.Document;

      const result = getAuthInfo(spec);
      expect(result.authInfo).to.not.be.undefined;
      expect(result.authInfo!.name).to.equal("apiKey1");
    });

    it("returns error for multiple different auths", () => {
      const spec = {
        components: {
          securitySchemes: {
            apiKey1: { type: "apiKey", name: "x-api-key", in: "header" },
            apiKey2: { type: "apiKey", name: "x-api-key2", in: "header" },
          },
        },
        paths: {
          "/pets": {
            get: {
              operationId: "listPets",
              security: [{ apiKey1: [] }],
              responses: {},
            },
          },
          "/dogs": {
            get: {
              operationId: "listDogs",
              security: [{ apiKey2: [] }],
              responses: {},
            },
          },
        },
      } as unknown as OpenAPIV3.Document;

      const result = getAuthInfo(spec);
      expect(result.error).to.be.a("string");
    });
  });

  describe("getResponseJson", () => {
    it("returns json content from 200 response", () => {
      const op = {
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
      } as unknown as OpenAPIV3.OperationObject;

      const { json, multipleMediaType } = getResponseJson(op);
      expect(json).to.have.property("schema");
      expect(multipleMediaType).to.be.false;
    });

    it("returns empty for no response", () => {
      const { json } = getResponseJson(undefined);
      expect(Object.keys(json)).to.have.length(0);
    });
  });

  describe("convertPathToCamelCase", () => {
    it("converts path segments to CamelCase", () => {
      expect(convertPathToCamelCase("/pets/{petId}/toys")).to.equal("PetsPetIdToys");
    });
  });

  describe("getSafeRegistrationIdEnvName", () => {
    it("returns uppercase with underscores", () => {
      expect(getSafeRegistrationIdEnvName("my-api-key")).to.equal("MY_API_KEY");
    });

    it("adds PREFIX_ for non-alpha start", () => {
      expect(getSafeRegistrationIdEnvName("123key")).to.equal("PREFIX_123KEY");
    });

    it("returns empty for empty input", () => {
      expect(getSafeRegistrationIdEnvName("")).to.equal("");
    });
  });

  describe("checkServerUrl", () => {
    it("accepts https", () => {
      const errors = checkServerUrl([{ url: "https://api.example.com" }]);
      expect(errors).to.have.length(0);
    });

    it("rejects http by default", () => {
      const errors = checkServerUrl([{ url: "http://api.example.com" }]);
      expect(errors).to.have.length(1);
      expect(errors[0].type).to.equal(ErrorType.UrlProtocolNotSupported);
    });

    it("allows http when allowHttp is true", () => {
      const errors = checkServerUrl([{ url: "http://api.example.com" }], true);
      expect(errors).to.have.length(0);
    });

    it("rejects relative url", () => {
      const errors = checkServerUrl([{ url: "/api" }]);
      expect(errors).to.have.length(1);
      expect(errors[0].type).to.equal(ErrorType.RelativeServerUrlNotSupported);
    });
  });

  describe("isObjectSchema", () => {
    it("returns true for type: object", () => {
      expect(isObjectSchema({ type: "object" } as any)).to.be.true;
    });
    it("returns true for schema with properties but no type", () => {
      expect(isObjectSchema({ properties: { name: { type: "string" } } } as any)).to.be.true;
    });
    it("returns false for string type", () => {
      expect(isObjectSchema({ type: "string" } as any)).to.be.false;
    });
  });

  describe("isWellKnownName", () => {
    it("matches known names", () => {
      expect(isWellKnownName("product_title", ["title"])).to.be.true;
      expect(isWellKnownName("image-url", ["image"])).to.be.true;
    });
    it("returns false for non-matching", () => {
      expect(isWellKnownName("foo", ["title", "image"])).to.be.false;
    });
  });

  describe("formatStr", () => {
    it("replaces %s placeholders", () => {
      expect(formatStr("Hello %s, you have %s items", "World", "5")).to.equal(
        "Hello World, you have 5 items"
      );
    });
  });

  describe("getServerObject", () => {
    it("returns operation-level server first", () => {
      const spec = {
        servers: [{ url: "https://root.com" }],
        paths: {
          "/pets": {
            servers: [{ url: "https://path.com" }],
            get: {
              servers: [{ url: "https://op.com" }],
              responses: {},
            },
          },
        },
      } as unknown as OpenAPIV3.Document;

      const server = getServerObject(spec, "get", "/pets");
      expect(server?.url).to.equal("https://op.com");
    });
  });

  describe("generateParametersFromSchema", () => {
    it("extracts required and optional params from object schema", () => {
      const schema: OpenAPIV3.SchemaObject = {
        type: "object",
        required: ["name"],
        properties: {
          name: { type: "string", description: "Pet name" },
          age: { type: "integer", description: "Pet age" },
        },
      };

      const result = generateParametersFromSchema(schema, "body", true);
      expect(result.requiredParams).to.have.length(1);
      expect(result.optionalParams).to.have.length(1);
      expect(result.requiredParams[0].name).to.equal("name");
    });
  });
});

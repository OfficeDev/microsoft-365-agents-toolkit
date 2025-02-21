// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert, expect } from "chai";
import "mocha";
import { Utils } from "../src/utils";
import { OpenAPIV3 } from "openapi-types";
import { ConstantString } from "../src/constants";
import { ErrorType, ProjectType, ParseOptions, AuthInfo } from "../src/interfaces";

describe("utils", () => {
  describe("updateFirstLetter", () => {
    it("should return the string with the first letter capitalized", () => {
      const result = Utils.updateFirstLetter("hello");
      expect(result).to.equal("Hello");
    });

    it("should return an empty string if the input is empty", () => {
      const result = Utils.updateFirstLetter("");
      expect(result).to.equal("");
    });
  });

  describe("isNotSupportedAuth", () => {
    it("should return false for an empty authSchemeArray", () => {
      const authSchemeArray: AuthInfo[][] = [];
      expect(Utils.isNotSupportedAuth(authSchemeArray)).to.equal(false);
    });

    it("should return true if all authSchemeArray elements have more than one auth", () => {
      const authSchemeArray: AuthInfo[][] = [
        [
          {
            name: "oauth",
            authScheme: {
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "https://example.com/oauth/authorize",
                  tokenUrl: "https://example.com/oauth/token",
                  scopes: {
                    read: "Grants read access",
                    write: "Grants write access",
                    admin: "Grants access to admin operations",
                  },
                },
              },
            },
          },
          {
            name: "apikey",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
        ],
      ];
      expect(Utils.isNotSupportedAuth(authSchemeArray)).to.equal(true);
    });

    it("should return false if any authSchemeArray element has a single OAuth with Auth Code Flow", () => {
      const authSchemeArray: AuthInfo[][] = [
        [
          {
            name: "oauth",
            authScheme: {
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "https://example.com/oauth/authorize",
                  tokenUrl: "https://example.com/oauth/token",
                  scopes: {
                    read: "Grants read access",
                    write: "Grants write access",
                    admin: "Grants access to admin operations",
                  },
                },
              },
            },
          },
        ],
      ];
      expect(Utils.isNotSupportedAuth(authSchemeArray)).to.equal(false);
    });

    it("should return false if any authSchemeArray element has a single Bearer Token Auth", () => {
      const authSchemeArray: AuthInfo[][] = [
        [
          {
            name: "apikey",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
        ],
      ];
      expect(Utils.isNotSupportedAuth(authSchemeArray)).to.equals(false);
    });

    it("should return true if all authSchemeArray elements have unsupported auth schemes", () => {
      const authSchemeArray: AuthInfo[][] = [
        [
          {
            authScheme: {
              type: "http" as const,
              scheme: "basic",
            },
            name: "basic_auth1",
          },
        ],
        [
          {
            authScheme: {
              type: "http" as const,
              scheme: "basic",
            },
            name: "basic_auth2",
          },
        ],
      ];
      expect(Utils.isNotSupportedAuth(authSchemeArray)).to.equals(true);
    });
  });

  describe("isObjectSchema", () => {
    it('should return true when schema.type is "object"', () => {
      const schema: OpenAPIV3.SchemaObject = { type: "object" };
      expect(Utils.isObjectSchema(schema)).to.be.true;
    });

    it("should return true when schema.type is not defined but schema.properties is defined", () => {
      const schema: OpenAPIV3.SchemaObject = { properties: { prop1: { type: "string" } } };
      expect(Utils.isObjectSchema(schema)).to.be.true;
    });

    it("should return false when schema.type is not defined and schema.properties is not defined", () => {
      const schema: OpenAPIV3.SchemaObject = {};
      expect(Utils.isObjectSchema(schema)).to.be.false;
    });

    it('should return false when schema.type is defined but not "object"', () => {
      const schema: OpenAPIV3.SchemaObject = { type: "string" };
      expect(Utils.isObjectSchema(schema)).to.be.false;
    });
  });

  describe("convertPathToCamelCase", () => {
    it("should convert a path to camel case", () => {
      const path = "this/is/a/{test}/path";
      const expected = "ThisIsATestPath";
      const result = Utils.convertPathToCamelCase(path);
      assert.strictEqual(result, expected);
    });

    it("should convert a path to camel case start with /", () => {
      const path = "/this/is/a/{test}/path";
      const expected = "ThisIsATestPath";
      const result = Utils.convertPathToCamelCase(path);
      assert.strictEqual(result, expected);
    });

    it("should return an empty string for an empty path", () => {
      const path = "";
      const expected = "";
      const result = Utils.convertPathToCamelCase(path);
      assert.strictEqual(result, expected);
    });

    it("should return the same string for a path with no slashes", () => {
      const path = "test";
      const expected = "Test";
      const result = Utils.convertPathToCamelCase(path);
      assert.strictEqual(result, expected);
    });

    it("should return correct result for string with {} and .", () => {
      const path = "/{section}.json";
      const expected = "SectionJson";
      const result = Utils.convertPathToCamelCase(path);
      assert.strictEqual(result, expected);
    });

    it("should return correct result for complex string", () => {
      const path = "/{section}.{test1}/{test2}.json";
      const expected = "SectionTest1Test2Json";
      const result = Utils.convertPathToCamelCase(path);
      assert.strictEqual(result, expected);
    });
  });

  describe("getUrlProtocol", () => {
    it("should return the protocol of a valid URL", () => {
      const url = "https://example.com/path/to/file";
      const protocol = Utils.getUrlProtocol(url);
      expect(protocol).to.equal("https:");
    });

    it("should return undefined for an invalid URL", () => {
      const url = "not a url";
      const protocol = Utils.getUrlProtocol(url);
      expect(protocol).to.be.undefined;
    });

    it("should return undefined for relative url", () => {
      const url = "/v3";
      const protocol = Utils.getUrlProtocol(url);
      expect(protocol).to.be.undefined;
    });

    it("should return the protocol for other protocol", () => {
      const url = "ftp://v1";
      const protocol = Utils.getUrlProtocol(url);
      expect(protocol).to.equal("ftp:");
    });
  });

  describe("checkServerUrl", () => {
    it("should return an empty array if the server URL is valid", () => {
      const servers = [{ url: "https://example.com" }];
      const errors = Utils.checkServerUrl(servers);
      assert.deepStrictEqual(errors, []);
    });

    it("should return an error if the server URL is relative", () => {
      const servers = [{ url: "/api" }];
      const errors = Utils.checkServerUrl(servers);
      assert.deepStrictEqual(errors, [
        {
          type: ErrorType.RelativeServerUrlNotSupported,
          content: ConstantString.RelativeServerUrlNotSupported,
          data: servers,
        },
      ]);
    });

    it("should not return an error if the server URL is http and allowHttp is true", () => {
      const servers = [{ url: "http://example.com" }];
      const errors = Utils.checkServerUrl(servers, true);
      assert.deepStrictEqual(errors, []);
    });

    it("should return an error if the server URL protocol is not HTTPS", () => {
      const servers = [{ url: "http://example.com" }];
      const errors = Utils.checkServerUrl(servers);
      assert.deepStrictEqual(errors, [
        {
          type: ErrorType.UrlProtocolNotSupported,
          content: Utils.format(ConstantString.UrlProtocolNotSupported, "http"),
          data: "http",
        },
      ]);
    });
  });

  describe("validateServer", () => {
    it("should return an error if there is no server information", () => {
      const spec = { paths: {} };

      const options: ParseOptions = {
        allowMissingId: true,
        allowAPIKeyAuth: false,
        allowMultipleParameters: false,
        allowOauth2: false,
        projectType: ProjectType.SME,
        allowMethods: ["get", "post"],
      };

      const errors = Utils.validateServer(spec as OpenAPIV3.Document, options);
      assert.deepStrictEqual(errors, [
        {
          type: ErrorType.NoServerInformation,
          content: ConstantString.NoServerInformation,
        },
      ]);
    });

    it("should return an error if protocol is not supported ", () => {
      const spec = {
        paths: {
          "/api": {
            get: {
              servers: [{ url: "ftp://example.com" }],
            },
          },
        },
      };

      const options: ParseOptions = {
        allowMissingId: true,
        allowAPIKeyAuth: false,
        allowMultipleParameters: false,
        allowOauth2: false,
        projectType: ProjectType.SME,
        allowMethods: ["get", "post"],
      };

      const errors = Utils.validateServer(spec as any, options);
      assert.deepStrictEqual(errors, [
        {
          type: ErrorType.UrlProtocolNotSupported,
          content: Utils.format(ConstantString.UrlProtocolNotSupported, "ftp"),
          data: "ftp",
        },
      ]);
    });

    it("should validate top-level servers", () => {
      const spec = {
        servers: [{ url: "https://example.com" }],
        paths: {},
      };

      const options: ParseOptions = {
        allowMissingId: true,
        allowAPIKeyAuth: false,
        allowMultipleParameters: false,
        allowOauth2: false,
        projectType: ProjectType.SME,
        allowMethods: ["get", "post"],
      };

      const errors = Utils.validateServer(spec as OpenAPIV3.Document, options);
      assert.deepStrictEqual(errors, []);
    });

    it("should validate path-level servers", () => {
      const spec = {
        paths: {
          "/api": {
            servers: [{ url: "https://example.com" }],
          },
        },
      };

      const options: ParseOptions = {
        allowMissingId: true,
        allowAPIKeyAuth: false,
        allowMultipleParameters: false,
        allowOauth2: false,
        projectType: ProjectType.SME,
        allowMethods: ["get", "post"],
      };

      const errors = Utils.validateServer(spec as any, options);
      assert.deepStrictEqual(errors, []);
    });

    it("should validate operation-level servers", () => {
      const spec = {
        paths: {
          "/api": {
            get: {
              servers: [{ url: "https://example.com" }],
              parameters: [
                {
                  in: "query",
                  schema: { type: "string" },
                },
              ],
              responses: {
                200: {
                  description: "OK",
                  content: {
                    "application/json": {
                      schema: {
                        type: "string",
                      },
                    },
                  },
                },
              },
            },
          },
        },
      };

      const options: ParseOptions = {
        allowMissingId: true,
        allowAPIKeyAuth: false,
        allowMultipleParameters: false,
        allowOauth2: false,
        projectType: ProjectType.SME,
        allowMethods: ["get", "post"],
      };

      const errors = Utils.validateServer(spec as any, options);
      assert.deepStrictEqual(errors, []);
    });

    it("should validate all levels of servers", () => {
      const spec = {
        servers: [{ url: "https://example.com" }],
        paths: {
          "/api": {
            servers: [{ url: "https://example.com" }],
            get: {
              parameters: [
                {
                  in: "query",
                  schema: { type: "string" },
                },
              ],
              servers: [{ url: "https://example.com" }],
              responses: {
                200: {
                  description: "OK",
                  content: {
                    "application/json": {
                      schema: {
                        type: "string",
                      },
                    },
                  },
                },
              },
            },
          },
        },
      };
      const options: ParseOptions = {
        allowMissingId: true,
        allowAPIKeyAuth: false,
        allowMultipleParameters: false,
        allowOauth2: false,
        projectType: ProjectType.SME,
        allowMethods: ["get", "post"],
      };

      const errors = Utils.validateServer(spec as any, options);
      assert.deepStrictEqual(errors, []);
    });

    it("should validate invalid server URLs", () => {
      const spec = {
        servers: [{ url: "/api" }],
        paths: {
          "/api": {
            servers: [{ url: "http://example.com" }],
            get: {
              parameters: [
                {
                  in: "query",
                  schema: { type: "string" },
                },
              ],
              servers: [{ url: "ftp://example.com" }],
              responses: {
                200: {
                  description: "OK",
                  content: {
                    "application/json": {
                      schema: {
                        type: "string",
                      },
                    },
                  },
                },
              },
            },
          },
        },
      };

      const options: ParseOptions = {
        allowMissingId: true,
        allowAPIKeyAuth: false,
        allowMultipleParameters: false,
        allowOauth2: false,
        projectType: ProjectType.SME,
        allowMethods: ["get", "post"],
      };

      const errors = Utils.validateServer(spec as any, options);
      assert.deepStrictEqual(errors, [
        {
          type: ErrorType.RelativeServerUrlNotSupported,
          content: ConstantString.RelativeServerUrlNotSupported,
          data: spec.servers,
        },
        {
          type: ErrorType.UrlProtocolNotSupported,
          content: Utils.format(ConstantString.UrlProtocolNotSupported, "http"),
          data: "http",
        },
        {
          type: ErrorType.UrlProtocolNotSupported,
          content: Utils.format(ConstantString.UrlProtocolNotSupported, "ftp"),
          data: "ftp",
        },
      ]);
    });
  });

  describe("getResponseJson", () => {
    it("should return an empty object if no JSON response is defined", () => {
      const operationObject = {};
      const { json, multipleMediaType } = Utils.getResponseJson(operationObject);
      expect(json).to.deep.equal({});
      expect(multipleMediaType).to.be.false;
    });

    it("should return an empty object if operationObject is undefined", () => {
      const operationObject = undefined;
      const { json, multipleMediaType } = Utils.getResponseJson(operationObject);
      expect(json).to.deep.equal({});
      expect(multipleMediaType).to.be.false;
    });

    it("should return the JSON response for status code 200", () => {
      const operationObject = {
        responses: {
          "200": {
            content: {
              "application/json": {
                schema: {
                  type: "object",
                  properties: {
                    message: { type: "string" },
                  },
                },
              },
            },
          },
        },
      } as any;
      const { json, multipleMediaType } = Utils.getResponseJson(operationObject);
      expect(json).to.deep.equal({
        schema: {
          type: "object",
          properties: {
            message: { type: "string" },
          },
        },
      });
      expect(multipleMediaType).to.be.false;
    });

    it("should return the JSON response for application/json; charset=utf-8;", () => {
      const operationObject = {
        responses: {
          "200": {
            content: {
              "application/json; charset=utf-8": {
                schema: {
                  type: "object",
                  properties: {
                    message: { type: "string" },
                  },
                },
              },
            },
          },
        },
      } as any;
      const { json, multipleMediaType } = Utils.getResponseJson(operationObject);
      expect(json).to.deep.equal({
        schema: {
          type: "object",
          properties: {
            message: { type: "string" },
          },
        },
      });
      expect(multipleMediaType).to.be.false;
    });

    it("should return empty JSON response for status code 200 with multiple media type", () => {
      const operationObject = {
        responses: {
          "200": {
            content: {
              "application/json": {
                schema: {
                  type: "object",
                  properties: {
                    message: { type: "string" },
                  },
                },
              },
              "application/xml": {
                schema: {
                  type: "object",
                  properties: {
                    message: { type: "string" },
                  },
                },
              },
            },
          },
        },
      } as any;
      const { json, multipleMediaType } = Utils.getResponseJson(operationObject);
      expect(json).to.deep.equal({});
      expect(multipleMediaType).to.be.true;
    });

    it("should return the JSON response for status code 201", () => {
      const operationObject = {
        responses: {
          "201": {
            content: {
              "application/json": {
                schema: {
                  type: "object",
                  properties: {
                    id: { type: "number" },
                  },
                },
              },
            },
          },
        },
      } as any;
      const { json, multipleMediaType } = Utils.getResponseJson(operationObject);
      expect(json).to.deep.equal({
        schema: {
          type: "object",
          properties: {
            id: { type: "number" },
          },
        },
      });

      expect(multipleMediaType).to.be.false;
    });

    it("should return the JSON response for the default status code", () => {
      const operationObject = {
        responses: {
          default: {
            content: {
              "application/json": {
                schema: {
                  type: "object",
                  properties: {
                    error: { type: "string" },
                  },
                },
              },
            },
          },
        },
      } as any;
      const { json, multipleMediaType } = Utils.getResponseJson(operationObject);
      expect(json).to.deep.equal({
        schema: {
          type: "object",
          properties: {
            error: { type: "string" },
          },
        },
      });
      expect(multipleMediaType).to.be.false;
    });

    it("should return the JSON response for the 200 status code", () => {
      const operationObject = {
        responses: {
          "201": {
            content: {
              "application/json": {
                schema: {
                  type: "object",
                  properties: {
                    id: { type: "number" },
                  },
                },
              },
            },
          },
          "200": {
            content: {
              "application/json": {
                schema: {
                  type: "object",
                  properties: {
                    message: { type: "string" },
                  },
                },
              },
            },
          },
        },
      } as any;
      const { json, multipleMediaType } = Utils.getResponseJson(operationObject);
      expect(json).to.deep.equal({
        schema: {
          type: "object",
          properties: {
            message: { type: "string" },
          },
        },
      });
      expect(multipleMediaType).to.be.false;
    });

    it("should return an empty object if all JSON responses are undefined", () => {
      const operationObject = {
        responses: {
          "400": {
            content: {
              "application/xml": {
                schema: {
                  type: "object",
                  properties: {
                    error: { type: "string" },
                  },
                },
              },
            },
          },
          "500": {
            content: {
              "text/plain": {
                schema: {
                  type: "object",
                  properties: {
                    message: { type: "string" },
                  },
                },
              },
            },
          },
        },
      } as any;
      const { json, multipleMediaType } = Utils.getResponseJson(operationObject);
      expect(json).to.deep.equal({});
      expect(multipleMediaType).to.be.false;
    });
  });

  describe("resolveServerUrl", () => {
    it("should replace one environment variables in the URL", () => {
      process.env.OPENAPI_SERVER_URL = "https://localhost:3000/api";
      const url = "${{OPENAPI_SERVER_URL}}";
      const expectedUrl = "https://localhost:3000/api";
      const resolvedUrl = Utils.resolveEnv(url);
      assert.strictEqual(resolvedUrl, expectedUrl);
    });

    it("should throw an error if environment variable is not defined", () => {
      delete process.env.OPENAPI_SERVER_URL;
      const url = "${{OPENAPI_SERVER_URL}}";
      const expectedUrl = "https://localhost:3000/api";
      assert.throws(
        () => Utils.resolveEnv(url),
        Error,
        Utils.format(ConstantString.ResolveServerUrlFailed, "OPENAPI_SERVER_URL")
      );
    });

    it("should replace multiple environment variables in the URL", () => {
      process.env.API_HOST = "localhost";
      process.env.API_PORT = "3000";
      const url = "http://${{API_HOST}}:${{API_PORT}}/api";
      const expectedUrl = "http://localhost:3000/api";
      const resolvedUrl = Utils.resolveEnv(url);
      assert.strictEqual(resolvedUrl, expectedUrl);
    });

    it("should throw an error if one environment variable is not defined", () => {
      delete process.env.API_PORT;
      process.env.API_HOST = "localhost";
      const url = "http://${{API_HOST}}:${{API_PORT}}/api";
      assert.throws(
        () => Utils.resolveEnv(url),
        Error,
        Utils.format(ConstantString.ResolveServerUrlFailed, "API_PORT")
      );
    });
  });

  describe("isWellKnownName", () => {
    it("should return true for well-known result property names", () => {
      expect(Utils.isWellKnownName("result", ConstantString.WellknownResultNames)).to.be.true;
      expect(Utils.isWellKnownName("r_e_s_u_l_t", ConstantString.WellknownResultNames)).to.be.true;
      expect(Utils.isWellKnownName("r-e-s-u-l-t", ConstantString.WellknownResultNames)).to.be.true;
      expect(Utils.isWellKnownName("data", ConstantString.WellknownResultNames)).to.be.true;
      expect(Utils.isWellKnownName("items", ConstantString.WellknownResultNames)).to.be.true;
      expect(Utils.isWellKnownName("root", ConstantString.WellknownResultNames)).to.be.true;
      expect(Utils.isWellKnownName("matches", ConstantString.WellknownResultNames)).to.be.true;
      expect(Utils.isWellKnownName("queries", ConstantString.WellknownResultNames)).to.be.true;
      expect(Utils.isWellKnownName("list", ConstantString.WellknownResultNames)).to.be.true;
      expect(Utils.isWellKnownName("output", ConstantString.WellknownResultNames)).to.be.true;
    });

    it("should return true for well-known result property names with different casing", () => {
      expect(Utils.isWellKnownName("Result", ConstantString.WellknownResultNames)).to.be.true;
      expect(Utils.isWellKnownName("DaTa", ConstantString.WellknownResultNames)).to.be.true;
      expect(Utils.isWellKnownName("ITEMS", ConstantString.WellknownResultNames)).to.be.true;
      expect(Utils.isWellKnownName("Root", ConstantString.WellknownResultNames)).to.be.true;
      expect(Utils.isWellKnownName("MaTcHeS", ConstantString.WellknownResultNames)).to.be.true;
      expect(Utils.isWellKnownName("QuErIeS", ConstantString.WellknownResultNames)).to.be.true;
      expect(Utils.isWellKnownName("LiSt", ConstantString.WellknownResultNames)).to.be.true;
      expect(Utils.isWellKnownName("OutPut", ConstantString.WellknownResultNames)).to.be.true;
    });

    it("should return true for name substring is well-known result property names", () => {
      expect(Utils.isWellKnownName("testResult", ConstantString.WellknownResultNames)).to.be.true;
      expect(Utils.isWellKnownName("carData", ConstantString.WellknownResultNames)).to.be.true;
      expect(Utils.isWellKnownName("productItems", ConstantString.WellknownResultNames)).to.be.true;
      expect(Utils.isWellKnownName("rootValue", ConstantString.WellknownResultNames)).to.be.true;
      expect(Utils.isWellKnownName("matchesResult", ConstantString.WellknownResultNames)).to.be
        .true;
      expect(Utils.isWellKnownName("DataQueries", ConstantString.WellknownResultNames)).to.be.true;
      expect(Utils.isWellKnownName("productLists", ConstantString.WellknownResultNames)).to.be.true;
      expect(Utils.isWellKnownName("outputData", ConstantString.WellknownResultNames)).to.be.true;
    });

    it("should return false for non well-known result property names", () => {
      expect(Utils.isWellKnownName("foo", ConstantString.WellknownResultNames)).to.be.false;
      expect(Utils.isWellKnownName("bar", ConstantString.WellknownResultNames)).to.be.false;
      expect(Utils.isWellKnownName("baz", ConstantString.WellknownResultNames)).to.be.false;
      expect(Utils.isWellKnownName("qux", ConstantString.WellknownResultNames)).to.be.false;
    });
  });

  describe("format", () => {
    it("should replace %s placeholders with arguments", () => {
      const result = Utils.format("Hello, %s!", "world");
      expect(result).to.equal("Hello, world!");
    });

    it("should handle multiple placeholders and arguments", () => {
      const result = Utils.format("The %s is %s.", "answer", "42");
      expect(result).to.equal("The answer is 42.");
    });

    it("should handle missing arguments", () => {
      const result = Utils.format("Hello, %s!", "");
      expect(result).to.equal("Hello, !");
    });

    it("should handle extra arguments", () => {
      const result = Utils.format("Hello, %s!", "world", "extra");
      expect(result).to.equal("Hello, world!");
    });

    it("should handle no placeholders", () => {
      const result = Utils.format("Hello, world!");
      expect(result).to.equal("Hello, world!");
    });
  });

  describe("getSafeRegistrationIdEnvName", () => {
    it("should return an empty string if authName is not provided", () => {
      expect(Utils.getSafeRegistrationIdEnvName("")).to.equal("");
    });

    it("should replace non-alphanumeric characters with underscores and convert to uppercase", () => {
      expect(Utils.getSafeRegistrationIdEnvName("auth@name")).to.equal("AUTH_NAME");
    });

    it('should prefix the result with "PREFIX_" if it does not start with an uppercase letter', () => {
      expect(Utils.getSafeRegistrationIdEnvName("1authname")).to.equal("PREFIX_1AUTHNAME");
    });

    it('should not prefix the result with "PREFIX_" if it starts with an uppercase letter', () => {
      expect(Utils.getSafeRegistrationIdEnvName("Authname")).to.equal("AUTHNAME");
    });

    it("should convert all lowercase letters to uppercase", () => {
      expect(Utils.getSafeRegistrationIdEnvName("authname")).to.equal("AUTHNAME");
    });
  });

  describe("getAuthSchemaObject", () => {
    it("should return correct oauth scheme with refresh url", () => {
      const authParameter = {
        apis: "operationId1",
        authorizationUrl: "mockedAuthorizationUrl",
        tokenUrl: "mockedTokenUrl",
        refreshUrl: "mockedRefreshUrl",
        scopes: {
          read: "Grants read access",
        },
      };
      const authSchema = Utils.getAuthSchemaObject("oauth", authParameter);
      expect(authSchema).to.deep.equal({
        type: "oauth2",
        flows: {
          authorizationCode: {
            authorizationUrl: "mockedAuthorizationUrl",
            tokenUrl: "mockedTokenUrl",
            refreshUrl: "mockedRefreshUrl",
            scopes: {
              read: "Grants read access",
            },
          },
        },
      });
    });

    it("should return correct oauth scheme with refreshout url", () => {
      const authParameter = {
        apis: "operationId1",
        authorizationUrl: "mockedAuthorizationUrl",
        tokenUrl: "mockedTokenUrl",
        refreshUrl: undefined,
        scopes: {
          read: "Grants read access",
        },
      };
      const authSchema = Utils.getAuthSchemaObject("oauth", authParameter);
      expect(authSchema).to.deep.equal({
        type: "oauth2",
        flows: {
          authorizationCode: {
            authorizationUrl: "mockedAuthorizationUrl",
            tokenUrl: "mockedTokenUrl",
            refreshUrl: undefined,
            scopes: {
              read: "Grants read access",
            },
          },
        },
      });
    });

    it("should return correct apiKey scheme", () => {
      const authParameter = {
        apis: "operationId1",
        in: "header",
        name: "apiKeyName",
      };
      const authSchema = Utils.getAuthSchemaObject("api-key", authParameter);
      expect(authSchema).to.deep.equal({
        type: "apiKey",
        in: "header",
        name: "apiKeyName",
      });
    });

    it("should return correct bearer token scheme", () => {
      const authParameter = {
        apis: "operationId1",
      };
      const authSchema = Utils.getAuthSchemaObject("bearer-token", authParameter);
      expect(authSchema).to.deep.equal({
        type: "http",
        scheme: "bearer",
      });
    });
  });
});

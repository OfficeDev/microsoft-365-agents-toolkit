// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";
import { expect } from "chai";
import * as sinon from "sinon";
import { SpecParser } from "../../src/specParser.browser";
import { ErrorType, ValidationStatus, WarningType } from "../../src/interfaces";
import { SpecParserError } from "../../src/specParserError";
import { ConstantString } from "../../src/constants";
import { OpenAPIV3 } from "openapi-types";
import { Utils } from "../../src/utils";
import SwaggerParser from "@apidevtools/swagger-parser";
import { SMEValidator } from "../../src/validators/smeValidator";

describe("SpecParser in Browser", () => {
  afterEach(() => {
    sinon.restore();
  });

  describe("listSupporttedAPIInfo", () => {
    it("should list parameters successfully with required and optional parameters", async () => {
      const specPath = "valid-spec.yaml";
      const specParser = new SpecParser(specPath, {
        allowMissingId: false,
        allowMultipleParameters: true,
      });
      const spec = {
        openapi: "3.0.0",
        info: {
          title: "Repair Service",
          description: "A simple service to manage repairs",
          version: "1.0.0",
        },
        servers: [
          {
            url: "https://poc-apim-gateway-fkh0bdaufkfpdugz.b02.azurefd.net",
            description: "The repair api server",
          },
        ],
        paths: {
          "/assignRepair": {
            get: {
              operationId: "assignRepair",
              summary:
                "Assign repair to technician for the customer based on car type and repair type",
              description:
                "Assign repair to technician for the customer based on car type and repair type",
              parameters: [
                {
                  name: "carType",
                  in: "query",
                  description: "Car type to repair",
                  schema: {
                    type: "string",
                  },
                  required: false,
                },
                {
                  name: "customerStatus",
                  in: "query",
                  description: "Customer status",
                  schema: {
                    type: "string",
                    enum: ["available", "pending", "sold"],
                  },
                },
                {
                  name: "customerToggle",
                  in: "query",
                  description: "Customer Toggle",
                  schema: {
                    type: "boolean",
                  },
                  allowEmptyValue: true,
                  required: true,
                },
                {
                  name: "limit",
                  in: "query",
                  schema: {
                    type: "integer",
                    minimum: 1,
                    maximum: 50,
                  },
                  description: "The numbers of items to return",
                  required: true,
                },
                {
                  name: "startDateString",
                  in: "query",
                  schema: {
                    type: "string",
                  },
                  description: "The start date for the report in different format",
                  required: true,
                },
              ],
              responses: {
                "200": {
                  description: "The response that represents an appointment for the repair",
                  content: {
                    "application/json": {
                      schema: {
                        type: "object",
                        properties: {
                          id: {
                            type: "string",
                            description: "Id of the repair",
                          },
                          title: {
                            type: "string",
                            description: "The short summary of the repair",
                          },
                          assignedTo: {
                            type: "string",
                            description: "The engineer who is responsible for the repair",
                          },
                          customerPhoneNumber: {
                            type: "string",
                            description: "The phone number of the customer",
                          },
                          date: {
                            type: "string",
                            format: "date-time",
                            description:
                              "The date and time when the repair is scheduled or completed",
                          },
                          image: {
                            type: "string",
                            format: "uri",
                            description:
                              "The URL of the image of the item to be repaired or the repair process",
                          },
                        },
                      },
                    },
                  },
                },
              },
            },
          },
        },
      };

      const parseStub = sinon.stub(specParser.parser, "parse").resolves(spec as any);
      const dereferenceStub = sinon.stub(specParser.parser, "dereference").resolves(spec as any);

      const result = await specParser.listSupportedAPIInfo();
      expect(result).to.deep.equal([
        {
          method: "GET",
          path: "/assignRepair",
          title: "Assign repair to technician for ",
          id: "assignRepair",
          description:
            "Assign repair to technician for the customer based on car type and repair type",
          parameters: [
            {
              name: "customerToggle",
              title: "CustomerToggle",
              description: "Customer Toggle",
              inputType: "toggle",
              isRequired: true,
            },
            {
              name: "limit",
              title: "Limit",
              description: "The numbers of items to return",
              inputType: "number",
              isRequired: true,
            },
            {
              name: "startDateString",
              title: "StartDateString",
              description: "The start date for the report in different format",
              inputType: "text",
              isRequired: true,
            },
            {
              name: "carType",
              title: "CarType",
              description: "Car type to repair",
              inputType: "text",
            },
            {
              name: "customerStatus",
              title: "CustomerStatus",
              description: "Customer status",
              inputType: "choiceset",
              choices: [
                {
                  title: "available",
                  value: "available",
                },
                {
                  title: "pending",
                  value: "pending",
                },
                {
                  title: "sold",
                  value: "sold",
                },
              ],
            },
          ],
        },
      ]);
    });

    it("should return a list of HTTP methods and paths for all GET with 1 parameter and without security", async () => {
      const specPath = "valid-spec.yaml";
      const specParser = new SpecParser(specPath, { allowMissingId: false });
      const spec = {
        servers: [
          {
            url: "https://example.com",
          },
        ],
        paths: {
          "/pets": {
            get: {
              operationId: "getPetById",
              security: [{ api_key: [] }],
            },
          },
          "/user/{userId}": {
            get: {
              operationId: "getUserById",
              description: "Get user by user id, balabala",
              summary: "Get user by user id",
              parameters: [
                {
                  name: "userId",
                  in: "path",
                  description: "User Id",
                  schema: {
                    type: "string",
                  },
                },
              ],
              responses: {
                200: {
                  content: {
                    "application/json": {
                      schema: {
                        type: "object",
                        properties: {
                          name: {
                            type: "string",
                          },
                        },
                      },
                    },
                  },
                },
              },
            },
            post: {
              operationId: "createUser",
              security: [{ api_key: [] }],
            },
          },
          "/store/order": {
            post: {
              operationId: "placeOrder",
            },
          },
        },
      };

      const parseStub = sinon.stub(specParser.parser, "parse").resolves(spec as any);
      const dereferenceStub = sinon.stub(specParser.parser, "dereference").resolves(spec as any);

      const result = await specParser.listSupportedAPIInfo();
      expect(result).to.deep.equal([
        {
          method: "GET",
          path: "/user/{userId}",
          title: "Get user by user id",
          id: "getUserById",
          parameters: [
            {
              name: "userId",
              title: "UserId",
              description: "User Id",
            },
          ],
          description: "Get user by user id, balabala",
        },
      ]);
    });

    it("should contain warning for GET with 2 optional parameter and without security", async () => {
      const specPath = "valid-spec.yaml";
      const specParser = new SpecParser(specPath, { allowMissingId: false });
      const spec = {
        servers: [
          {
            url: "https://example.com",
          },
        ],
        paths: {
          "/user/{userId}": {
            get: {
              operationId: "getUserById",
              description: "Get user by user id, balabala",
              summary: "Get user by user id",
              parameters: [
                {
                  name: "userId",
                  in: "path",
                  description: "User Id",
                  schema: {
                    type: "string",
                  },
                },
                {
                  name: "name",
                  in: "query",
                  description: "User Name",
                  schema: {
                    type: "string",
                  },
                },
              ],
              responses: {
                200: {
                  content: {
                    "application/json": {
                      schema: {
                        type: "object",
                        properties: {
                          name: {
                            type: "string",
                          },
                        },
                      },
                    },
                  },
                },
              },
            },
            post: {
              operationId: "createUser",
              security: [{ api_key: [] }],
            },
          },
          "/store/order": {
            post: {
              operationId: "placeOrder",
            },
          },
        },
      };

      const parseStub = sinon.stub(specParser.parser, "parse").resolves(spec as any);
      const dereferenceStub = sinon.stub(specParser.parser, "dereference").resolves(spec as any);

      const result = await specParser.listSupportedAPIInfo();
      expect(result).to.deep.equal([
        {
          method: "GET",
          path: "/user/{userId}",
          title: "Get user by user id",
          id: "getUserById",
          parameters: [
            {
              name: "userId",
              title: "UserId",
              description: "User Id",
            },
            {
              name: "name",
              title: "Name",
              description: "User Name",
            },
          ],
          description: "Get user by user id, balabala",
        },
      ]);
    });

    it("should not list api without operationId with allowMissingId is true", async () => {
      const specPath = "valid-spec.yaml";
      const specParser = new SpecParser(specPath, { allowMissingId: true });
      const spec = {
        servers: [
          {
            url: "https://example.com",
          },
        ],
        paths: {
          "/pets": {
            get: {
              operationId: "getPetById",
              security: [{ api_key: [] }],
            },
          },
          "/user/{userId}": {
            get: {
              parameters: [
                {
                  name: "userId",
                  in: "path",
                  schema: {
                    type: "string",
                  },
                },
              ],
              responses: {
                200: {
                  content: {
                    "application/json": {
                      schema: {
                        type: "object",
                        properties: {
                          name: {
                            type: "string",
                          },
                        },
                      },
                    },
                  },
                },
              },
            },
            post: {
              operationId: "createUser",
              security: [{ api_key: [] }],
            },
          },
          "/store/order": {
            post: {
              operationId: "placeOrder",
            },
          },
        },
      };

      const parseStub = sinon.stub(specParser.parser, "parse").resolves(spec as any);
      const dereferenceStub = sinon.stub(specParser.parser, "dereference").resolves(spec as any);

      const result = await specParser.listSupportedAPIInfo();

      expect(result).to.deep.equal([]);
    });

    it("should throw an error when the SwaggerParser library throws an error", async () => {
      const specPath = "invalid-spec.yaml";
      const specParser = new SpecParser(specPath);
      sinon.stub(SwaggerParser.prototype, "validate").rejects(new Error("Invalid specification"));
      const parseStub = sinon
        .stub(specParser.parser, "parse")
        .rejects(new Error("Invalid specification"));
      try {
        await specParser.listSupportedAPIInfo();
        expect.fail("Expected an error to be thrown");
      } catch (err) {
        expect((err as SpecParserError).message).contain("Invalid specification");
        expect((err as SpecParserError).errorType).to.equal(ErrorType.listSupportedAPIInfoFailed);
      }
    });
  });

  describe("validate", () => {
    it("should return an error result when the spec is not valid", async () => {
      const specParser = new SpecParser("/path/to/spec.yaml");
      const spec = { openapi: "3.0.0" };
      sinon.stub(specParser.parser, "parse").resolves(spec as any);
      sinon.stub(specParser.parser, "dereference").resolves(spec as any);
      const parseStub = sinon
        .stub(specParser.parser, "validate")
        .rejects(new Error("Invalid spec"));

      const result = await specParser.validate();

      expect(result.status).to.equal(ValidationStatus.Error);
      expect(result.warnings).to.be.an("array").that.is.empty;
      expect(result.errors).to.be.an("array").that.has.lengthOf(1);
      expect(result.errors[0].type).to.equal(ErrorType.SpecNotValid);
      expect(result.errors[0].content).to.equal("Error: Invalid spec");
      sinon.assert.calledOnce(parseStub);
    });

    it("should return error object if the spec version is 2.0 with allowSwagger is false", async function () {
      const specPath = "path/to/spec";
      const spec = {
        swagger: "2.0",
        info: {
          version: "1.0.0",
          title: "Swagger Petstore",
          description:
            "A sample API that uses a petstore as an example to demonstrate features in the swagger-2.0 specification",
        },
        host: "petstore.swagger.io",
        basePath: "/v2",
        schemes: ["https"],
        paths: {
          "/pet": {
            post: {
              summary: "Add a new pet to the store",
              operationId: "addPet",
              consumes: ["application/json"],
              produces: ["application/json"],
              parameters: [
                {
                  in: "body",
                  name: "body",
                  schema: {
                    type: "object",
                    required: ["name"],
                    properties: {
                      id: {
                        type: "integer",
                        format: "int64",
                      },
                      name: {
                        type: "string",
                      },
                    },
                  },
                },
              ],
              responses: {
                "200": {
                  description: "Pet added to the store",
                  schema: {
                    type: "object",
                    required: ["id", "name"],
                    properties: {
                      id: {
                        type: "integer",
                        format: "int64",
                      },
                      name: {
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
      const specParser = new SpecParser(specPath, { allowSwagger: false });

      const parseStub = sinon.stub(specParser.parser, "parse").resolves(spec as any);

      const dereferenceStub = sinon.stub(specParser.parser, "dereference").resolves(spec);
      const validateStub = sinon.stub(specParser.parser, "validate").resolves(spec);

      const result = await specParser.validate();

      expect(result).to.deep.equal({
        status: ValidationStatus.Error,
        errors: [
          {
            type: ErrorType.SwaggerNotSupported,
            content: ConstantString.SwaggerNotSupported,
          },
        ],
        warnings: [],
      });
    });

    it("should return an error result object if no server information", async function () {
      const specPath = "path/to/spec";
      const spec = { openapi: "3.0.0" };

      const specParser = new SpecParser(specPath);
      const parseStub = sinon.stub(specParser.parser, "parse").resolves(spec as any);
      const dereferenceStub = sinon.stub(specParser.parser, "dereference").resolves(spec as any);
      const validateStub = sinon.stub(specParser.parser, "validate").resolves(spec as any);
      const result = await specParser.validate();

      expect(result).to.deep.equal({
        status: ValidationStatus.Error,
        warnings: [],
        errors: [
          { type: ErrorType.NoServerInformation, content: ConstantString.NoServerInformation },
          { type: ErrorType.NoSupportedApi, content: ConstantString.NoSupportedApi, data: [] },
        ],
      });
      sinon.assert.calledOnce(dereferenceStub);
    });

    it("should return no supported API error with invalid api info", async function () {
      const specPath = "path/to/spec";
      const spec = {
        openapi: "3.0.2",
        servers: [
          {
            url: "https://servers1",
          },
        ],
        paths: {
          "/pet": {
            get: {
              tags: ["pet"],
              summary: "Get pet information from the store",
              parameters: [
                {
                  name: "tags",
                  in: "query",
                  description: "Tags to filter by",
                  schema: {
                    type: "string",
                  },
                },
              ],
              responses: {
                "200": {
                  content: {
                    "application/json": {
                      schema: {
                        $ref: "#/components/schemas/Pet",
                      },
                    },
                  },
                },
              },
            },
          },
        },
      };

      const specParser = new SpecParser(specPath, { allowMissingId: false });
      const parseStub = sinon.stub(specParser.parser, "parse").resolves(spec as any);
      const dereferenceStub = sinon.stub(specParser.parser, "dereference").resolves(spec as any);
      const validateStub = sinon.stub(specParser.parser, "validate").resolves(spec as any);
      const result = await specParser.validate();

      expect(result).to.deep.equal({
        status: ValidationStatus.Error,
        warnings: [],
        errors: [
          {
            type: ErrorType.NoSupportedApi,
            content: ConstantString.NoSupportedApi,
            data: [
              {
                api: "GET /pet",
                reason: [ErrorType.MissingOperationId],
              },
            ],
          },
        ],
      });
      sinon.assert.calledOnce(dereferenceStub);
    });

    it("should return an error result object if server url is http", async function () {
      const specPath = "path/to/spec";
      const spec = { openapi: "3.0.0", servers: [{ url: "http://server1" }] };

      const specParser = new SpecParser(specPath);
      const parseStub = sinon.stub(specParser.parser, "parse").resolves(spec as any);
      const dereferenceStub = sinon.stub(specParser.parser, "dereference").resolves(spec as any);
      const validateStub = sinon.stub(specParser.parser, "validate").resolves(spec as any);
      const result = await specParser.validate();

      expect(result).to.deep.equal({
        status: ValidationStatus.Error,
        warnings: [],
        errors: [
          {
            type: ErrorType.UrlProtocolNotSupported,
            content: Utils.format(ConstantString.UrlProtocolNotSupported, "http"),
            data: "http",
          },
          { type: ErrorType.NoSupportedApi, content: ConstantString.NoSupportedApi, data: [] },
        ],
      });
      sinon.assert.calledOnce(dereferenceStub);
    });

    it("should return an error result object if server url is relative path", async function () {
      const specPath = "path/to/spec";
      const spec = { openapi: "3.0.0", servers: [{ url: "path/to/server1" }] };

      const specParser = new SpecParser(specPath);
      const parseStub = sinon.stub(specParser.parser, "parse").resolves(spec as any);
      const dereferenceStub = sinon.stub(specParser.parser, "dereference").resolves(spec as any);
      const validateStub = sinon.stub(specParser.parser, "validate").resolves(spec as any);
      const result = await specParser.validate();

      expect(result).to.deep.equal({
        status: ValidationStatus.Error,
        warnings: [],
        errors: [
          {
            type: ErrorType.RelativeServerUrlNotSupported,
            content: ConstantString.RelativeServerUrlNotSupported,
            data: [
              {
                url: "path/to/server1",
              },
            ],
          },
          { type: ErrorType.NoSupportedApi, content: ConstantString.NoSupportedApi, data: [] },
        ],
      });
      sinon.assert.calledOnce(dereferenceStub);
    });

    it("should return an error result object if no supported apis", async function () {
      const specPath = "path/to/spec";
      const spec = { openapi: "3.0.0", servers: [{ url: "https://server1" }] };

      const specParser = new SpecParser(specPath);
      const parseStub = sinon.stub(specParser.parser, "parse").resolves(spec as any);
      const dereferenceStub = sinon.stub(specParser.parser, "dereference").resolves(spec as any);
      const validateStub = sinon.stub(specParser.parser, "validate").resolves(spec as any);
      const result = await specParser.validate();

      expect(result).to.deep.equal({
        status: ValidationStatus.Error,
        warnings: [],
        errors: [
          { type: ErrorType.NoSupportedApi, content: ConstantString.NoSupportedApi, data: [] },
        ],
      });
      sinon.assert.calledOnce(dereferenceStub);
    });

    it("should return an error result object if contain remote reference", async function () {
      const spec = {
        openapi: "3.0.3",
        info: {
          title: "Swagger Petstore - OpenAPI 3.",
          version: "1.0.11",
        },
        servers: [
          {
            url: "https://petstore3.swagger.io/api/v3",
          },
        ],
        paths: {
          "/pet": {
            get: {
              tags: ["pet"],
              operationId: "updatePet",
              responses: {
                "200": {
                  description: "Successful operation",
                  content: {
                    "application/json": {
                      schema: {
                        $ref: "https://petstore3.swagger.io/api/v3/openapi.json#/components/schemas/Pet",
                      },
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
                name: {
                  type: "string",
                  example: "doggie",
                },
              },
            },
          },
        },
      } as OpenAPIV3.Document;
      const specParser = new SpecParser(spec as any);
      sinon.stub(specParser.parser, "parse").resolves(spec as any);
      sinon.stub(specParser.parser, "dereference").resolves(spec as any);
      const validateStub = sinon.stub(specParser.parser, "validate").resolves(spec as any);
      sinon
        .stub(specParser.parser.$refs, "paths")
        .returns(["localPath", "https://petstore3.swagger.io/api/v3/openapi.json"]);
      const result = await specParser.validate();

      expect(result.errors[0].type).equal(ErrorType.RemoteRefNotSupported);
      expect(result.status).equal(ValidationStatus.Error);
    });

    it("should return an warning result object if missing operation id", async function () {
      const specPath = "path/to/spec";
      const spec = {
        openapi: "3.0.2",
        servers: [
          {
            url: "https://servers1",
          },
        ],
        paths: {
          "/pet": {
            get: {
              tags: ["pet"],
              summary: "Get pet information from the store",
              parameters: [
                {
                  name: "tags",
                  in: "query",
                  description: "Tags to filter by",
                  schema: {
                    type: "string",
                  },
                },
              ],
              responses: {
                "200": {
                  content: {
                    "application/json": {
                      schema: {
                        $ref: "#/components/schemas/Pet",
                      },
                    },
                  },
                },
              },
            },
          },
        },
      };

      const specParser = new SpecParser(specPath, { allowMissingId: true });
      const parseStub = sinon.stub(specParser.parser, "parse").resolves(spec as any);
      const dereferenceStub = sinon.stub(specParser.parser, "dereference").resolves(spec as any);
      const validateStub = sinon.stub(specParser.parser, "validate").resolves(spec as any);
      const result = await specParser.validate();

      expect(result).to.deep.equal({
        status: ValidationStatus.Warning,
        warnings: [
          {
            type: WarningType.OperationIdMissing,
            content: Utils.format(ConstantString.MissingOperationId, "GET /pet"),
            data: ["GET /pet"],
          },
        ],
        errors: [],
      });
      sinon.assert.calledOnce(dereferenceStub);
    });

    it("should return a valid result when the spec is valid", async () => {
      const specPath = "path/to/spec";
      const spec = {
        openapi: "3.0.2",
        servers: [
          {
            url: "https://server1",
          },
        ],
        paths: {
          "/pet": {
            get: {
              tags: ["pet"],
              operationId: "getPet",
              summary: "Get pet information from the store",
              parameters: [
                {
                  name: "tags",
                  in: "query",
                  description: "Tags to filter by",
                  schema: {
                    type: "string",
                  },
                },
              ],
              responses: {
                "200": {
                  content: {
                    "application/json": {
                      schema: {
                        $ref: "#/components/schemas/Pet",
                      },
                    },
                  },
                },
              },
            },
          },
        },
      };

      const specParser = new SpecParser(specPath);
      const parseStub = sinon.stub(specParser.parser, "parse").resolves(spec as any);
      const dereferenceStub = sinon.stub(specParser.parser, "dereference").resolves(spec as any);
      const validateStub = sinon.stub(specParser.parser, "validate").resolves(spec as any);
      const result = await specParser.validate();
      expect(result.status).to.equal(ValidationStatus.Valid);
      expect(result.warnings).to.be.an("array").that.is.empty;
      expect(result.errors).to.be.an("array").that.is.empty;
      sinon.assert.calledOnce(dereferenceStub);
    });

    it("should throw a SpecParserError when an error occurs", async () => {
      const specPath = "path/to/spec";
      const spec = {
        openapi: "3.0.2",
        servers: [
          {
            url: "https://server1",
          },
        ],
        paths: {
          "/pet": {
            get: {
              tags: ["pet"],
              operationId: "getPet",
              summary: "Get pet information from the store",
              parameters: [
                {
                  name: "tags",
                  in: "query",
                  description: "Tags to filter by",
                  schema: {
                    type: "string",
                  },
                },
              ],
              responses: {
                "200": {
                  content: {
                    "application/xml": {
                      schema: {
                        $ref: "#/components/schemas/Pet",
                      },
                    },
                  },
                },
              },
            },
          },
        },
      };

      try {
        const specParser = new SpecParser(specPath);
        const parseStub = sinon.stub(specParser.parser, "parse").resolves(spec as any);
        const dereferenceStub = sinon.stub(specParser.parser, "dereference").resolves(spec as any);
        const validateStub = sinon.stub(specParser.parser, "validate").resolves(spec as any);
        sinon.stub(SMEValidator.prototype, "validateSpec").throws(new Error("validateSpec error"));

        const result = await specParser.validate();
        expect.fail("Expected SpecParserError to be thrown");
      } catch (err: any) {
        expect(err).to.be.instanceOf(SpecParserError);
        expect(err.errorType).to.equal(ErrorType.ValidateFailed);
        expect(err.message).to.equal("Error: validateSpec error");
      }
    });
  });

  describe("generate", () => {
    it("should throw not implemented error", async () => {
      try {
        const specParser = new SpecParser("path/to/spec.yaml");
        const filter = ["get /hello"];
        const outputSpecPath = "path/to/output.yaml";
        const result = await specParser.generate(
          "path/to/manifest.json",
          filter,
          outputSpecPath,
          "path/to/adaptiveCardFolder"
        );
        expect.fail("Should throw not implemented error");
      } catch (error: any) {
        expect(error.message).to.equal("Method not implemented.");
      }
    });
  });

  describe("generateForCopilot", () => {
    it("should throw not implemented error", async () => {
      try {
        const specParser = new SpecParser("path/to/spec.yaml");
        const filter = ["get /hello"];
        const outputSpecPath = "path/to/output.yaml";
        const result = await specParser.generateForCopilot(
          "path/to/manifest.json",
          filter,
          outputSpecPath,
          "ai-plugin"
        );
        expect.fail("Should throw not implemented error");
      } catch (error: any) {
        expect(error.message).to.equal("Method not implemented.");
      }
    });
  });

  describe("list", () => {
    it("should throw an error when the SwaggerParser library throws an error", async () => {
      try {
        const specPath = "valid-spec.yaml";
        const specParser = new SpecParser(specPath);
        await specParser.list();
        expect.fail("Should throw not implemented error");
      } catch (error: any) {
        expect(error.message).to.equal("Method not implemented.");
      }
    });
  });
});

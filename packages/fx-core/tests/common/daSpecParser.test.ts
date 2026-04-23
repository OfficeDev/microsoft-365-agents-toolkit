// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import "mocha";
import sinon from "sinon";
import * as fs from "fs-extra";
import * as os from "os";
import * as path from "path";
import * as kiotaClient from "../../src/common/kiotaClient";
import * as daSpecParser from "../../src/common/daSpecParser";
import * as utils from "../../src/common/utils";
import { featureFlagManager, FeatureFlags } from "../../src/common/featureFlags";
import { Platform } from "@microsoft/teamsfx-api";
import { KiotaOpenApiNode, KiotaTreeResult, OpenApiSpecVersion } from "@microsoft/kiota";
import {
  AdaptiveCardUpdateStrategy,
  ErrorType,
  ValidationStatus,
  WarningResult,
  WarningType,
} from "@microsoft/m365-spec-parser";

describe("daSpecParser", () => {
  let listAPITreeInfoStub: sinon.SinonStub;
  let featureFlagStub: sinon.SinonStub;
  let isJsonSpecFileStub: sinon.SinonStub;
  let parseAndUpdatePluginManifestStub: sinon.SinonStub;

  beforeEach(() => {
    listAPITreeInfoStub = sinon.stub(kiotaClient, "listAPITreeInfo");
    featureFlagStub = sinon.stub(featureFlagManager, "getBooleanValue");
    isJsonSpecFileStub = sinon.stub(utils, "isJsonSpecFile");
    parseAndUpdatePluginManifestStub = sinon.stub(
      daSpecParser,
      "parseAndUpdatePluginManifestForKiota"
    );
    parseAndUpdatePluginManifestStub.callsFake(async (pluginPath, updatePlaceholder) => {
      // This ensures we don't actually call the real implementation
      return [];
    });
    featureFlagStub.withArgs(FeatureFlags.KiotaNPMIntegration).returns(true);
    isJsonSpecFileStub.resolves(false);
  });

  afterEach(() => {
    sinon.restore();
  });

  describe("listAPIInfo with KiotaNPMIntegration enabled", () => {
    it("should return empty result when treeInfo is {}", async () => {
      listAPITreeInfoStub.resolves({});

      const result = await daSpecParser.listAPIInfo("path/to/spec");

      assert.deepEqual(result, {
        specVersion: undefined,
        allAPICount: 0,
        validAPICount: 0,
        APIs: [],
      });
    });

    it("should return empty result when rootNode is undefined", async () => {
      listAPITreeInfoStub.resolves({ rootNode: undefined });

      const result = await daSpecParser.listAPIInfo("path/to/spec");

      assert.deepEqual(result, {
        specVersion: undefined,
        allAPICount: 0,
        validAPICount: 0,
        APIs: [],
      });
    });

    it("should extract operations with simple node structure", async () => {
      const mockTreeInfo: KiotaTreeResult = {
        rootNode: {
          isOperation: true,
          path: "api/resource#GET",
          segment: "GET",
          operationId: "getResource",
          summary: "Get resource",
          description: "Get a specific resource",
          selected: true,
          children: [],
        } as KiotaOpenApiNode,
        servers: ["https://api.example.com"],
        security: [],
        securitySchemes: {},
        logs: [],
        specVersion: OpenApiSpecVersion.V3_0,
      };

      listAPITreeInfoStub.resolves(mockTreeInfo);

      const result = await daSpecParser.listAPIInfo("path/to/spec");

      assert.equal(result.allAPICount, 1);
      assert.equal(result.validAPICount, 1);
      assert.equal(result.APIs.length, 1);
      assert.deepEqual(result.APIs[0], {
        api: "GET api/resource",
        server: "https://api.example.com",
        operationId: "getResource",
        isValid: true,
        reason: [],
        auth: undefined,
        summary: "Get resource",
        description: "Get a specific resource",
      });
    });

    it("should extract operations with multiple # in path", async () => {
      const mockTreeInfo: KiotaTreeResult = {
        rootNode: {
          isOperation: true,
          path: "/#api/resource#GET",
          segment: "GET",
          operationId: "getResource",
          summary: "Get resource",
          description: "Get a specific resource",
          selected: true,
          children: [],
        } as KiotaOpenApiNode,
        servers: ["https://api.example.com"],
        security: [],
        securitySchemes: {},
        logs: [],
        specVersion: OpenApiSpecVersion.V3_0,
      };

      listAPITreeInfoStub.resolves(mockTreeInfo);

      const result = await daSpecParser.listAPIInfo("path/to/spec");

      assert.equal(result.allAPICount, 1);
      assert.equal(result.validAPICount, 1);
      assert.equal(result.APIs.length, 1);
      assert.deepEqual(result.APIs[0], {
        api: "GET /#api/resource",
        server: "https://api.example.com",
        operationId: "getResource",
        isValid: true,
        reason: [],
        auth: undefined,
        summary: "Get resource",
        description: "Get a specific resource",
      });
    });

    it("should not extract operations when not selected", async () => {
      const mockTreeInfo: KiotaTreeResult = {
        rootNode: {
          isOperation: true,
          path: "api/resource",
          segment: "GET",
          operationId: "getResource",
          summary: "Get resource",
          description: "Get a specific resource",
          selected: false,
          children: [],
        } as KiotaOpenApiNode,
        servers: ["https://api.example.com"],
        security: [],
        securitySchemes: {},
        logs: [],
        specVersion: OpenApiSpecVersion.V3_0,
      };

      listAPITreeInfoStub.resolves(mockTreeInfo);

      const result = await daSpecParser.listAPIInfo("path/to/spec");

      assert.equal(result.allAPICount, 0);
      assert.equal(result.validAPICount, 0);
      assert.equal(result.APIs.length, 0);
    });

    it("should handle Windows-style paths with backslashes", async () => {
      const mockTreeInfo: KiotaTreeResult = {
        rootNode: {
          isOperation: true,
          path: "api\\resource#GET",
          segment: "GET",
          operationId: "getResource",
          summary: "Get resource",
          description: "Get a specific resource",
          selected: true,
          children: [],
        } as KiotaOpenApiNode,
        servers: ["https://api.example.com"],
        security: [],
        securitySchemes: {},
        logs: [],
        specVersion: OpenApiSpecVersion.V3_0,
      };

      listAPITreeInfoStub.resolves(mockTreeInfo);

      const result = await daSpecParser.listAPIInfo("path/to/spec");

      assert.equal(result.APIs.length, 1);
      assert.equal(result.APIs[0].api, "GET api/resource");
    });

    it("should extract operations with child nodes", async () => {
      const mockTreeInfo: KiotaTreeResult = {
        rootNode: {
          isOperation: false,
          path: "api",
          segment: "",
          children: [
            {
              isOperation: true,
              path: "api/users#GET",
              segment: "GET",
              operationId: "getUsers",
              summary: "Get users",
              description: "Get all users",
              selected: true,
              children: [],
            },
            {
              isOperation: true,
              path: "api/posts#POST",
              segment: "POST",
              operationId: "createPost",
              summary: "Create post",
              description: "Create a new post",
              selected: true,
              children: [],
            },
          ],
        } as KiotaOpenApiNode,
        servers: ["https://api.example.com"],
        security: [],
        securitySchemes: {},
        logs: [],
        specVersion: OpenApiSpecVersion.V3_0,
      };

      listAPITreeInfoStub.resolves(mockTreeInfo);

      const result = await daSpecParser.listAPIInfo("path/to/spec");

      assert.equal(result.allAPICount, 2);
      assert.equal(result.validAPICount, 2);
      assert.equal(result.APIs.length, 2);
      assert.equal(result.APIs[0].api, "GET api/users");
      assert.equal(result.APIs[1].api, "POST api/posts");
    });

    it("should extract nested operations with security information", async () => {
      const mockTreeInfo: KiotaTreeResult = {
        rootNode: {
          isOperation: false,
          path: "api",
          segment: "",
          children: [
            {
              isOperation: false,
              path: "api/users",
              segment: "",
              children: [
                {
                  isOperation: true,
                  path: "api/users/profile#GET",
                  segment: "GET",
                  operationId: "getUserProfile",
                  summary: "Get profile",
                  description: "Get user profile",
                  selected: true,
                  children: [],
                },
              ],
            },
          ],
        } as KiotaOpenApiNode,
        logs: [],
        servers: ["https://api.example.com"],
        security: [{ api_key: [] }],
        securitySchemes: {
          api_key: {
            type: "apiKey",
            name: "x-api-key",
            in: "header",
            referenceId: "",
          },
        },
        specVersion: OpenApiSpecVersion.V3_0,
      };

      listAPITreeInfoStub.resolves(mockTreeInfo);

      const result = await daSpecParser.listAPIInfo("path/to/spec");

      assert.equal(result.allAPICount, 1);
      assert.equal(result.validAPICount, 1);
      assert.equal(result.APIs[0].api, "GET api/users/profile");
      assert.deepEqual(result.APIs[0].auth, {
        name: "api_key",
        authScheme: {
          type: "apiKey",
          name: "x-api-key",
          in: "header",
          referenceId: "",
        } as any,
      });
    });

    it("should use node-level servers and security if available", async () => {
      const mockTreeInfo: KiotaTreeResult = {
        rootNode: {
          isOperation: true,
          path: "api/resource",
          segment: "GET",
          operationId: "getResource",
          summary: "Get resource",
          description: "Get a specific resource",
          servers: ["https://node.example.com"],
          security: [{ node_auth: [] }],
          selected: true,
          children: [],
        } as KiotaOpenApiNode,
        servers: ["https://root.example.com"],
        security: [{ root_auth: [] }],
        securitySchemes: {
          node_auth: {
            type: "http",
            scheme: "bearer",
            referenceId: "",
          },
          root_auth: {
            type: "oauth2",
            flows: {},
            referenceId: "",
          },
        },
        logs: [],
        specVersion: OpenApiSpecVersion.V3_0,
      };

      listAPITreeInfoStub.resolves(mockTreeInfo);

      const result = await daSpecParser.listAPIInfo("path/to/spec");

      assert.equal(result.APIs[0].server, "https://node.example.com");
      assert.deepEqual(result.APIs[0].auth, {
        name: "node_auth",
        authScheme: {
          type: "http",
          scheme: "bearer",
          referenceId: "",
        } as any,
      });
    });

    it("should handle multiple security requirements", async () => {
      const mockTreeInfo: KiotaTreeResult = {
        rootNode: {
          isOperation: true,
          path: "api/secure",
          segment: "GET",
          operationId: "getSecureResource",
          summary: "Get secure resource",
          description: "Get a secure resource",
          selected: true,
          children: [],
        } as KiotaOpenApiNode,
        servers: ["https://api.example.com"],
        security: [{ api_key: [], oauth2: [] }],
        securitySchemes: {
          api_key: {
            type: "apiKey",
            name: "x-api-key",
            in: "header",
            referenceId: "",
          },
          oauth2: {
            type: "oauth2",
            flows: {},
            referenceId: "",
          },
        },
        logs: [],
        specVersion: OpenApiSpecVersion.V3_0,
      };

      listAPITreeInfoStub.resolves(mockTreeInfo);

      const result = await daSpecParser.listAPIInfo("path/to/spec");

      assert.equal(result.APIs[0].auth?.name, "api_key, oauth2");
      assert.deepEqual(result.APIs[0].auth?.authScheme, {
        type: "multipleAuth",
      });
    });

    it("should handle securitySchemes is undefined", async () => {
      const mockTreeInfo: KiotaTreeResult = {
        rootNode: {
          isOperation: true,
          path: "api/secure",
          segment: "GET",
          operationId: "getSecureResource",
          summary: "Get secure resource",
          description: "Get a secure resource",
          selected: true,
          children: [],
        } as KiotaOpenApiNode,
        servers: ["https://api.example.com"],
        security: [{ api_key: [] }],
        securitySchemes: undefined,
        logs: [],
        specVersion: OpenApiSpecVersion.V3_0,
      };

      listAPITreeInfoStub.resolves(mockTreeInfo);

      const result = await daSpecParser.listAPIInfo("path/to/spec");

      assert.equal(result.APIs[0].auth, undefined);
    });

    it("should handle missing summary and description", async () => {
      const mockTreeInfo: KiotaTreeResult = {
        rootNode: {
          isOperation: true,
          path: "api/resource",
          segment: "GET",
          operationId: "getResource",
          selected: true,
          children: [],
        } as KiotaOpenApiNode,
        servers: ["https://api.example.com"],
        security: [],
        securitySchemes: {},
        logs: [],
        specVersion: OpenApiSpecVersion.V3_0,
      };

      listAPITreeInfoStub.resolves(mockTreeInfo);

      const result = await daSpecParser.listAPIInfo("path/to/spec");

      assert.equal(result.APIs[0].summary, "");
      assert.equal(result.APIs[0].description, "");
    });

    it("should properly handle platform parameter", async () => {
      const mockTreeInfo: KiotaTreeResult = {
        rootNode: {
          isOperation: true,
          path: "api/resource",
          segment: "GET",
          operationId: "getResource",
          summary: "Get resource",
          description: "Get a specific resource",
          selected: true,
          children: [],
        } as KiotaOpenApiNode,
        servers: ["https://api.example.com"],
        security: [],
        securitySchemes: {},
        logs: [],
        specVersion: OpenApiSpecVersion.V3_0,
      };

      listAPITreeInfoStub.resolves(mockTreeInfo);

      const result = await daSpecParser.listAPIInfo("path/to/spec", Platform.VS);

      assert.equal(result.allAPICount, 1);
      assert.equal(result.validAPICount, 1);
    });

    it("should handle undefined or empty security information", async () => {
      const mockTreeInfoNoSecurity: KiotaTreeResult = {
        rootNode: {
          isOperation: true,
          path: "api/resource",
          segment: "GET",
          operationId: "getResource",
          summary: "Get resource",
          description: "Get a specific resource",
          security: undefined,
          selected: true,
          children: [],
        } as KiotaOpenApiNode,
        servers: ["https://api.example.com"],
        security: undefined,
        securitySchemes: {},
        logs: [],
        specVersion: OpenApiSpecVersion.V3_0,
      };

      listAPITreeInfoStub.resolves(mockTreeInfoNoSecurity);
      const resultNoSecurity = await daSpecParser.listAPIInfo("path/to/spec");
      assert.isUndefined(resultNoSecurity.APIs[0].auth);

      const mockTreeInfoEmptySecurity: KiotaTreeResult = {
        rootNode: {
          isOperation: true,
          path: "api/resource",
          segment: "GET",
          operationId: "getResource",
          summary: "Get resource",
          description: "Get a specific resource",
          selected: true,
          children: [],
        } as KiotaOpenApiNode,
        servers: ["https://api.example.com"],
        security: [],
        securitySchemes: {},
        logs: [],
        specVersion: OpenApiSpecVersion.V3_0,
      };

      listAPITreeInfoStub.resolves(mockTreeInfoEmptySecurity);
      const resultEmptySecurity = await daSpecParser.listAPIInfo("path/to/spec");
      assert.isUndefined(resultEmptySecurity.APIs[0].auth);

      const mockTreeInfoEmptyRequirement: KiotaTreeResult = {
        rootNode: {
          isOperation: true,
          path: "api/resource",
          segment: "GET",
          operationId: "getResource",
          summary: "Get resource",
          description: "Get a specific resource",
          selected: true,
          children: [],
        } as KiotaOpenApiNode,
        servers: ["https://api.example.com"],
        security: [{}],
        securitySchemes: {},
        logs: [],
        specVersion: OpenApiSpecVersion.V3_0,
      };

      listAPITreeInfoStub.resolves(mockTreeInfoEmptyRequirement);
      const resultEmptyRequirement = await daSpecParser.listAPIInfo("path/to/spec");
      assert.isUndefined(resultEmptyRequirement.APIs[0].auth);
    });

    it("should validate server information and authentication types correctly", async () => {
      const mockTreeInfo: KiotaTreeResult = {
        rootNode: {
          isOperation: false,
          path: "api",
          segment: "",
          children: [
            {
              isOperation: true,
              path: "api/noserver",
              segment: "GET",
              operationId: "getNoServer",
              servers: [],
              selected: true,
              children: [],
            },
            {
              isOperation: true,
              path: "api/invalidserver",
              segment: "GET",
              operationId: "getInvalidServer",
              servers: ["example/index.html"],
              selected: true,
              children: [],
            },
            {
              isOperation: true,
              path: "api/multipleauth",
              segment: "GET",
              operationId: "getMultipleAuth",
              servers: ["https://valid.example.com"],
              security: [{ auth1: [], auth2: [] }],
              selected: true,
              children: [],
            },
            {
              isOperation: true,
              path: "api/apikey",
              segment: "GET",
              operationId: "getWithAPIKey",
              servers: ["https://valid.example.com"],
              security: [{ api_key_auth: [] }],
              selected: true,
              children: [],
            },
            {
              isOperation: true,
              path: "api/oauth",
              segment: "GET",
              operationId: "getWithOAuth",
              servers: ["https://valid.example.com"],
              security: [{ oauth_auth: [] }],
              selected: true,
              children: [],
            },
            {
              isOperation: true,
              path: "api/bearer",
              segment: "GET",
              operationId: "getWithBearer",
              servers: ["https://valid.example.com"],
              security: [{ bearer_auth: [] }],
              selected: true,
              children: [],
            },
          ],
        } as KiotaOpenApiNode,
        servers: [],
        security: [],
        securitySchemes: {
          auth1: { type: "apiKey", name: "key1", in: "header", referenceId: "" },
          auth2: { type: "oauth2", flows: {}, referenceId: "" },
          api_key_auth: { type: "apiKey", name: "x-api-key", in: "header", referenceId: "" },
          oauth_auth: {
            type: "oauth2",
            flows: {
              authorizationCode: {
                authorizationUrl: "https://example.com/auth",
                tokenUrl: "https://example.com/token",
                scopes: {},
              },
            },
            referenceId: "",
          },
          bearer_auth: { type: "http", scheme: "bearer", referenceId: "" },
        },
        logs: [],
        specVersion: OpenApiSpecVersion.V3_0,
      };

      listAPITreeInfoStub.resolves(mockTreeInfo);

      const resultNonVS = await daSpecParser.listAPIInfo("path/to/spec", Platform.VSCode);

      const noServerOp = resultNonVS.APIs.find((op) => op.operationId === "getNoServer");
      assert.isDefined(noServerOp);
      assert.isTrue(noServerOp!.reason.includes(ErrorType.NoServerInformation));

      const invalidServerOp = resultNonVS.APIs.find((op) => op.operationId === "getInvalidServer");
      assert.isDefined(invalidServerOp);
      assert.isTrue(invalidServerOp!.reason.includes(ErrorType.RelativeServerUrlNotSupported));
    });

    it("should handle undefined servers and securitySchemes in treeInfo", async () => {
      const mockTreeInfo: KiotaTreeResult = {
        rootNode: {
          isOperation: true,
          path: "api/resource",
          segment: "GET",
          operationId: "getResource",
          summary: "Get resource",
          description: "Get a specific resource",
          selected: true,
          children: [],
        } as KiotaOpenApiNode,
        security: [],
        logs: [],
        specVersion: OpenApiSpecVersion.V3_0,
      };

      listAPITreeInfoStub.resolves(mockTreeInfo);

      const result = await daSpecParser.listAPIInfo("path/to/spec");

      assert.equal(result.allAPICount, 1);
      assert.equal(result.APIs.length, 1);
      assert.isUndefined(result.APIs[0].server);
      assert.isUndefined(result.APIs[0].auth);
    });

    it("should handle undefined security in child node", async () => {
      const mockTreeInfo: KiotaTreeResult = {
        rootNode: {
          isOperation: false,
          path: "api",
          segment: "",
          children: [
            {
              isOperation: true,
              path: "api/resource",
              segment: "GET",
              operationId: "getResource",
              children: [],
              selected: true,
            },
          ],
        } as KiotaOpenApiNode,
        servers: ["https://api.example.com"],
        security: [{ api_key: [] }],
        securitySchemes: {
          api_key: { type: "apiKey", name: "x-api-key", in: "header", referenceId: "" },
        },
        logs: [],
        specVersion: OpenApiSpecVersion.V3_0,
      };

      listAPITreeInfoStub.resolves(mockTreeInfo);

      const result = await daSpecParser.listAPIInfo("path/to/spec");

      assert.equal(result.APIs.length, 1);
      assert.isDefined(result.APIs[0].auth);
      assert.equal(result.APIs[0].auth!.name, "api_key");
    });
    it("should specifically test multipleAuth type detection", async () => {
      const mockTreeInfo: KiotaTreeResult = {
        rootNode: {
          isOperation: true,
          path: "api/multi-auth",
          segment: "GET",
          operationId: "getWithMultiAuth",
          servers: ["https://valid.example.com"],
          security: [{ auth1: [], auth2: [] }],
          selected: true,
          children: [],
        } as KiotaOpenApiNode,
        servers: ["https://api.example.com"],
        security: [],
        securitySchemes: {
          auth1: { type: "apiKey", name: "x-api-key", in: "header", referenceId: "" },
          auth2: { type: "oauth2", flows: {}, referenceId: "" },
        },
        logs: [],
        specVersion: OpenApiSpecVersion.V3_0,
      };

      listAPITreeInfoStub.resolves(mockTreeInfo);

      const checkServerUrlStub = sinon.stub().returns([]);
      sinon.replace(
        require("@microsoft/m365-spec-parser").Utils,
        "checkServerUrl",
        checkServerUrlStub
      );

      const result = await daSpecParser.listAPIInfo("path/to/spec");

      assert.equal(result.APIs.length, 1);
      assert.equal(result.APIs[0].auth!.authScheme.type, "multipleAuth");
      assert.isTrue(result.APIs[0].isValid);
    });
  });

  describe("validateOpenAPISpec with KiotaNPMIntegration enabled", () => {
    it("should handle errors in listAPIInfo", async () => {
      const errorMessage = "Failed to parse spec";
      listAPITreeInfoStub.rejects(new Error(errorMessage));

      const result = await daSpecParser.validateOpenAPISpec("path/to/spec");

      assert.equal(result.status, ValidationStatus.Error);
      assert.deepEqual(result.errors, [
        {
          type: ErrorType.SpecNotValid,
          content: `OpenAPI specification file is not valid: Error: ${errorMessage}`,
        },
      ]);
      assert.equal(result.specHash, "");
    });

    it("should return error when no APIs found", async () => {
      const mockTreeInfo: KiotaTreeResult = {
        rootNode: {
          isOperation: false,
          path: "api",
          segment: "",
          children: [],
        } as KiotaOpenApiNode,
        servers: [],
        security: [],
        securitySchemes: {},
        logs: [],
        specVersion: OpenApiSpecVersion.V3_0,
      };

      listAPITreeInfoStub.resolves(mockTreeInfo);

      const result = await daSpecParser.validateOpenAPISpec("path/to/spec");

      assert.equal(result.status, ValidationStatus.Error);
      assert.deepEqual(result.errors, [{ type: ErrorType.NoSupportedApi, content: "", data: [] }]);
      assert.equal(result.specHash, "");
    });

    it("should return error when no valid APIs found", async () => {
      const mockTreeInfo: KiotaTreeResult = {
        rootNode: {
          isOperation: false,
          path: "api",
          segment: "",
          children: [
            {
              isOperation: true,
              path: "api/resource#GET",
              segment: "GET",
              operationId: "getResource",
              servers: [],
              selected: true,
              children: [],
            },
          ],
        } as KiotaOpenApiNode,
        servers: [],
        security: [],
        securitySchemes: {},
        logs: [],
        specVersion: OpenApiSpecVersion.V3_0,
      };
      listAPITreeInfoStub.resolves(mockTreeInfo);

      const checkServerUrlStub = sinon.stub();
      sinon.replace(
        require("@microsoft/m365-spec-parser").Utils,
        "checkServerUrl",
        checkServerUrlStub
      );
      checkServerUrlStub.returns([{ type: ErrorType.RelativeServerUrlNotSupported }]);

      const result = await daSpecParser.validateOpenAPISpec("path/to/spec");

      assert.equal(result.status, ValidationStatus.Error);
      assert.deepEqual(result.errors, [
        {
          type: ErrorType.NoSupportedApi,
          content: "",
          data: [
            {
              api: "GET api/resource",
              reason: ["no-server-information"],
            },
          ],
        },
      ]);
      assert.equal(result.specHash, "");
    });

    it("should return valid result with hash when valid APIs are found", async () => {
      const serverUrl = "https://api.example.com";
      const mockTreeInfo: KiotaTreeResult = {
        rootNode: {
          isOperation: true,
          path: "api/resource",
          segment: "GET",
          operationId: "getResource",
          servers: [serverUrl],
          selected: true,
          children: [],
        } as KiotaOpenApiNode,
        servers: [serverUrl],
        security: [],
        securitySchemes: {},
        logs: [],
        specVersion: OpenApiSpecVersion.V2_0,
      };

      listAPITreeInfoStub.resolves(mockTreeInfo);

      const checkServerUrlStub = sinon.stub();
      sinon.replace(
        require("@microsoft/m365-spec-parser").Utils,
        "checkServerUrl",
        checkServerUrlStub
      );
      checkServerUrlStub.returns([]);

      const result = await daSpecParser.validateOpenAPISpec("path/to/spec");

      assert.equal(result.status, ValidationStatus.Valid);
      assert.isEmpty(result.errors);
      assert.isTrue(result.warnings.length === 1);
      assert.isTrue(result.warnings[0].type === WarningType.ConvertSwaggerToOpenAPI);

      const expectedHash = require("crypto")
        .createHash("sha256")
        .update(JSON.stringify(serverUrl))
        .digest("hex");
      assert.equal(result.specHash, expectedHash);
    });

    it("should work when platform is VS", async () => {
      const serverUrl = "https://api.example.com";
      const mockTreeInfo: KiotaTreeResult = {
        rootNode: {
          isOperation: true,
          path: "api/resource",
          segment: "GET",
          operationId: "getResource",
          servers: [serverUrl],
          selected: true,
          children: [],
        } as KiotaOpenApiNode,
        servers: [serverUrl],
        security: [],
        securitySchemes: {},
        logs: [],
        specVersion: OpenApiSpecVersion.V3_1,
      };

      listAPITreeInfoStub.resolves(mockTreeInfo);

      const checkServerUrlStub = sinon.stub();
      sinon.replace(
        require("@microsoft/m365-spec-parser").Utils,
        "checkServerUrl",
        checkServerUrlStub
      );
      checkServerUrlStub.returns([]);

      const result = await daSpecParser.validateOpenAPISpec("path/to/spec", Platform.VS);
      assert.equal(result.status, ValidationStatus.Valid);
      assert.isTrue(result.warnings.length === 1);
      assert.isTrue(result.warnings[0].type === WarningType.OpenAPI31ConvertTo30);
    });
  });

  describe("generatePlugin with KiotaNPMIntegration enabled", () => {
    let kiotaGeneratePluginStub: sinon.SinonStub;
    let tmpDirSyncStub: sinon.SinonStub;
    let pathRelativeStub: sinon.SinonStub;

    beforeEach(() => {
      kiotaGeneratePluginStub = sinon.stub(kiotaClient, "kiotageneratePlugin");
      tmpDirSyncStub = sinon.stub(require("tmp"), "dirSync");
      pathRelativeStub = sinon.stub(require("path"), "relative");

      featureFlagStub.withArgs(FeatureFlags.KiotaNPMIntegration).returns(true);

      tmpDirSyncStub.returns({
        name: "c:\\tmp\\working-dir",
        removeCallback: sinon.stub(),
        unsafeCleanup: true,
      });
      kiotaGeneratePluginStub.resolves({
        openAPISpec: "c:\\tmp\\working-dir\\plugin\\openapi.yaml",
        aiPlugin: "c:\\tmp\\working-dir\\plugin\\ai-plugin.json",
        logs: [],
      });
      pathRelativeStub.returns("../openapi.yaml");
    });

    const pathMatcher = (expectedPath: string) =>
      sinon.match((actualPath) => {
        const normalizedActual = actualPath.replace(/\\/g, "/");
        const normalizedExpected = expectedPath.replace(/\\/g, "/");
        return normalizedActual === normalizedExpected;
      });

    it("should successfully generate plugin when feature flag is enabled", async () => {
      const specPath = "path/to/spec.yaml";
      const teamsManifestPath = "path/to/manifest.json";
      const outputAPISpecPath = "path/to/output/openapi.yaml";
      const outputAIPluginPath = "path/to/output/ai-plugin.json";
      const operations = ["GET /users", "POST /messages"];
      const adaptiveCardUpdateStrategy = AdaptiveCardUpdateStrategy.KeepExisting;

      const pathExistsStub = sinon.stub(require("fs-extra"), "pathExists").resolves(true);
      const readdirStub = sinon.stub(require("fs-extra"), "readdir").resolves(["openapi.json"]);
      const fsReadJSONStub = sinon
        .stub(require("fs-extra"), "readJSON")
        .resolves({ name: { short: "test-app" } });
      const fsWriteJsonStub = sinon.stub(require("fs-extra"), "writeJson").resolves();
      const fsCopyStub = sinon.stub(require("fs-extra"), "copy").resolves();

      const result = await daSpecParser.generatePlugin(
        specPath,
        teamsManifestPath,
        outputAPISpecPath,
        outputAIPluginPath,
        operations,
        adaptiveCardUpdateStrategy
      );

      assert.isTrue(tmpDirSyncStub.calledOnce);
      assert.isTrue(fsReadJSONStub.calledThrice);
      assert.isTrue(kiotaGeneratePluginStub.calledOnce);
      assert.deepEqual(kiotaGeneratePluginStub.firstCall.args[0], specPath);
      assert.deepEqual(
        kiotaGeneratePluginStub.firstCall.args[1].replace(/\\/g, "/"),
        "c:/tmp/working-dir/plugin"
      );
      assert.deepEqual(kiotaGeneratePluginStub.firstCall.args[2], "testapp");
      assert.deepEqual(kiotaGeneratePluginStub.firstCall.args[6], ["/users#GET", "/messages#POST"]);

      assert.equal(fsCopyStub.callCount, 3);

      const copyCallArgs = fsCopyStub.secondCall.args;
      assert.isTrue(copyCallArgs[0].replace(/\\/g, "/").endsWith("adaptiveCards"));
      assert.isTrue(copyCallArgs[0].replace(/\\/g, "/").endsWith("adaptiveCards"));

      assert.deepEqual(
        copyCallArgs[2],
        {
          overwrite: true,
          errorOnExist: false,
        },
        "Copy options don't match"
      );
      assert.isTrue(
        fsCopyStub.firstCall.calledWith(
          pathMatcher("c:/tmp/working-dir/plugin/openapi.yaml"),
          pathMatcher("path/to/output/openapi.yaml")
        )
      );

      assert.isTrue(
        fsCopyStub.thirdCall.calledWith(
          pathMatcher("c:/tmp/working-dir/.kiota/documents/testapp/openapi.json"),
          pathMatcher("path/to/output/openapi.yaml.original")
        )
      );

      assert.deepEqual(result, {
        allSuccess: true,
        warnings: [],
      });
    });

    it("should validate operations and generate appropriate warnings", async () => {
      const mockTreeInfo = {
        rootNode: {
          isOperation: false,
          path: "api",
          segment: "",
          children: [
            {
              isOperation: true,
              path: "api/missing-id",
              segment: "GET",
              summary: "Operation with missing ID",
              servers: ["https://valid.example.com"],
              selected: true,
              children: [],
            },
            {
              isOperation: true,
              path: "api/special-chars",
              segment: "POST",
              operationId: "create-resource",
              summary: "Operation with special characters in ID",
              servers: ["https://valid.example.com"],
              selected: true,
              children: [],
            },
            {
              isOperation: true,
              path: "api/unsupported-auth",
              segment: "GET",
              operationId: "getWithCustomAuth",
              summary: "Operation with unsupported auth",
              servers: ["https://valid.example.com"],
              security: [{ custom_auth: [] }],
              selected: true,
              children: [],
            },
          ],
        },
        servers: ["https://api.example.com"],
        security: [],
        securitySchemes: {
          custom_auth: { type: "http", scheme: "basic" },
        },
        logs: [],
      };

      listAPITreeInfoStub.resolves(mockTreeInfo);

      const specPath = "path/to/spec.json";
      const outputAPISpecPath = "path/to/output/openapi.spec";

      isJsonSpecFileStub.resolves(true);

      const readdirStub = sinon.stub(require("fs-extra"), "readdir").resolves(["openapi.json"]);
      const fsReadJSONStub = sinon
        .stub(require("fs-extra"), "readJSON")
        .resolves({ name: { short: "test-app" } });
      const fsCopyStub = sinon.stub(require("fs-extra"), "copy").resolves();
      const fsWriteJsonStub = sinon.stub(require("fs-extra"), "writeJson").resolves();

      const result = await daSpecParser.generatePlugin(
        specPath,
        "path/to/manifest.json",
        outputAPISpecPath,
        "path/to/output/ai-plugin.json",
        ["GET /api/missing-id", "POST /api/special-chars", "GET /api/unsupported-auth"],
        AdaptiveCardUpdateStrategy.KeepExisting
      );

      assert.isTrue(result.allSuccess);
      assert.equal(result.warnings.length, 3);
      assert.isTrue(
        result.warnings.some((w: WarningResult) => w.type === WarningType.OperationIdMissing)
      );
      assert.isTrue(
        result.warnings.some(
          (w: WarningResult) => w.type === WarningType.OperationIdContainsSpecialCharacters
        )
      );
      assert.isTrue(
        result.warnings.some((w: WarningResult) => w.type === WarningType.UnsupportedAuthType)
      );

      assert.isTrue(
        fsCopyStub.firstCall.calledWith(
          pathMatcher("c:/tmp/working-dir/plugin/openapi.yaml"),
          pathMatcher("path/to/output/openapi.yaml")
        )
      );
      assert.isTrue(
        fsCopyStub.secondCall.calledWith(
          pathMatcher("c:/tmp/working-dir/.kiota/documents/testapp/openapi.json"),
          pathMatcher("path/to/output/openapi.yaml.original")
        )
      );
    });

    it("should handle manifest with environment variables and special characters", async () => {
      const complexManifest = {
        name: {
          short: "Complex$App-Name_${{ENV_VAR}}",
        },
      };

      const readdirStub = sinon.stub(require("fs-extra"), "readdir").resolves(["openapi.json"]);
      const fsReadJSONStub = sinon.stub(require("fs-extra"), "readJSON").resolves(complexManifest);
      const fsCopyStub = sinon.stub(require("fs-extra"), "copy").resolves();
      const fsWriteJsonStub = sinon.stub(require("fs-extra"), "writeJson").resolves();

      await daSpecParser.generatePlugin(
        "path/to/spec.yaml",
        "path/to/manifest.json",
        "path/to/output/openapi.yaml",
        "path/to/output/ai-plugin.json",
        ["GET /api"],
        AdaptiveCardUpdateStrategy.KeepExisting
      );

      // Check namespace was properly sanitized
      assert.isTrue(kiotaGeneratePluginStub.calledOnce);
      // Instead of expecting just 'complexappname', allow for removal of vars
      const generatedNamespace = kiotaGeneratePluginStub.firstCall.args[2];
      assert.isString(generatedNamespace);
      assert.match(generatedNamespace, /^complexappname/);
    });

    it("should update plugin manifest with relative path", async () => {
      pathRelativeStub.returns("..\\..\\openapi.yaml");

      const readdirStub = sinon.stub(require("fs-extra"), "readdir").resolves(["openapi.json"]);
      const fsReadJSONStub = sinon.stub(require("fs-extra"), "readJSON").resolves({
        name: { short: "test-app" },
        runtimes: [
          {
            spec: { url: "old-path.yaml" },
          },
        ],
      });
      const fsCopyFileStub = sinon.stub(require("fs-extra"), "copy").resolves();
      const fsWriteJsonStub = sinon.stub(require("fs-extra"), "writeJson").resolves();

      await daSpecParser.generatePlugin(
        "path/to/spec.yaml",
        "path/to/manifest.json",
        "path/to/output/openapi.yaml",
        "path/to/output/ai-plugin.json",
        ["GET /api"],
        AdaptiveCardUpdateStrategy.KeepExisting
      );

      assert.isTrue(
        fsWriteJsonStub.calledWith(
          pathMatcher("path/to/output/ai-plugin.json"),
          sinon.match((value) => {
            return value.runtimes[0].spec.url === "../../openapi.yaml";
          }),
          { spaces: 4 }
        )
      );
    });

    it("should handle Windows paths and convert backslashes to forward slashes", async () => {
      pathRelativeStub.returns("..\\nested\\folder\\openapi.yaml");

      // Setup a mock plugin manifest with runtimes
      const pluginManifest = {
        name: { short: "test-app" },
        runtimes: [
          {
            spec: { url: "old-path.yaml" },
          },
        ],
      };

      const readdirStub = sinon.stub(require("fs-extra"), "readdir").resolves(["openapi.json"]);
      const fsReadJSONStub = sinon.stub(require("fs-extra"), "readJSON").resolves(pluginManifest);
      const fsCopyFileStub = sinon.stub(require("fs-extra"), "copy").resolves();
      const fsWriteJsonStub = sinon.stub(require("fs-extra"), "writeJson").resolves();

      await daSpecParser.generatePlugin(
        "path/to/spec.yaml",
        "path/to/manifest.json",
        "path/to/output/openapi.yaml",
        "path/to/output/ai-plugin.json",
        ["GET /api"],
        AdaptiveCardUpdateStrategy.KeepExisting
      );

      // Check that writeJson was called with the correct normalized path
      assert.isTrue(
        fsWriteJsonStub.calledWith(
          pathMatcher("path/to/output/ai-plugin.json"),
          sinon.match((value) => {
            return (
              value &&
              value.runtimes &&
              value.runtimes[0] &&
              value.runtimes[0].spec &&
              value.runtimes[0].spec.url === "../nested/folder/openapi.yaml"
            );
          }),
          { spaces: 4 }
        )
      );
    });

    it("should create correct include patterns from operations", async () => {
      const operations = [
        "GET /users",
        "POST /users",
        "PUT /users/{id}",
        "DELETE /users/{id}/comments",
        "PATCH /settings",
      ];

      const readdirStub = sinon.stub(require("fs-extra"), "readdir").resolves(["openapi.json"]);
      const fsReadJSONStub = sinon
        .stub(require("fs-extra"), "readJSON")
        .resolves({ name: { short: "test-app" } });
      const fsCopyFileStub = sinon.stub(require("fs-extra"), "copy").resolves();
      const fsWriteJsonStub = sinon.stub(require("fs-extra"), "writeJson").resolves();

      await daSpecParser.generatePlugin(
        "path/to/spec.yaml",
        "path/to/manifest.json",
        "path/to/output/openapi.yaml",
        "path/to/output/ai-plugin.json",
        operations,
        AdaptiveCardUpdateStrategy.KeepExisting
      );

      const expectedPatterns = [
        "/users#GET",
        "/users#POST",
        "/users/{id}#PUT",
        "/users/{id}/comments#DELETE",
        "/settings#PATCH",
      ];

      assert.deepEqual(kiotaGeneratePluginStub.firstCall.args[6], expectedPatterns);
    });

    it("should handle tree with completely missing optional fields", async () => {
      const mockTreeInfo = {
        rootNode: {
          isOperation: true,
          path: "api/resource",
          segment: "GET",
          operationId: "getResource",
          selected: true,
          children: [],
        },
        logs: [],
      };

      listAPITreeInfoStub.resolves(mockTreeInfo);

      const readdirStub = sinon.stub(require("fs-extra"), "readdir").resolves(["openapi.json"]);
      const fsReadJSONStub = sinon
        .stub(require("fs-extra"), "readJSON")
        .resolves({ name: { short: "test-app" } });
      const fsCopyFileStub = sinon.stub(require("fs-extra"), "copy").resolves();
      const fsWriteJsonStub = sinon.stub(require("fs-extra"), "writeJson").resolves();

      const result = await daSpecParser.generatePlugin(
        "path/to/spec.yaml",
        "path/to/manifest.json",
        "path/to/output/openapi.yaml",
        "path/to/output/ai-plugin.json",
        ["GET /api/resource"],
        AdaptiveCardUpdateStrategy.KeepExisting
      );

      assert.isTrue(result.allSuccess);
    });

    it("should handle both JSON and YAML original spec files", async () => {
      isJsonSpecFileStub.resolves(true);

      const readdirStub = sinon.stub(require("fs-extra"), "readdir").resolves(["openapi.json"]);
      const fsReadJSONStub = sinon
        .stub(require("fs-extra"), "readJSON")
        .resolves({ name: { short: "test-app" } });
      const fsCopyFileStub = sinon.stub(require("fs-extra"), "copy").resolves();
      const fsWriteJsonStub = sinon.stub(require("fs-extra"), "writeJson").resolves();

      await daSpecParser.generatePlugin(
        "path/to/spec.json",
        "path/to/manifest.json",
        "path/to/output/openapi.spec",
        "path/to/output/ai-plugin.json",
        ["GET /api"],
        AdaptiveCardUpdateStrategy.KeepExisting
      );

      assert.isTrue(
        fsCopyFileStub.secondCall.calledWith(
          pathMatcher("c:/tmp/working-dir/.kiota/documents/testapp/openapi.json"),
          pathMatcher("path/to/output/openapi.yaml.original")
        )
      );

      sinon.resetHistory();

      isJsonSpecFileStub.resolves(false);

      await daSpecParser.generatePlugin(
        "path/to/spec.yaml",
        "path/to/manifest.json",
        "path/to/output/openapi.spec",
        "path/to/output/ai-plugin.json",
        ["GET /api"],
        AdaptiveCardUpdateStrategy.KeepExisting
      );

      assert.isTrue(
        fsCopyFileStub.calledWith(
          pathMatcher("c:/tmp/working-dir/.kiota/documents/testapp/openapi.json"),
          pathMatcher("path/to/output/openapi.yaml.original")
        )
      );
    });

    it("should handle original spec file properly based on updateExistingPlugin flag", async () => {
      const readdirStub = sinon.stub(require("fs-extra"), "readdir").resolves(["openapi.json"]);
      const fsReadJSONStub = sinon.stub(require("fs-extra"), "readJSON");
      const fsCopyFileStub = sinon.stub(require("fs-extra"), "copy").resolves();
      const fsWriteJsonStub = sinon.stub(require("fs-extra"), "writeJson").resolves();

      fsReadJSONStub.resolves({
        name: { short: "test-app" },
        runtimes: [
          {
            spec: { url: "old-path.yaml" },
          },
        ],
      });

      await daSpecParser.generatePlugin(
        "path/to/spec.yaml",
        "path/to/manifest.json",
        "path/to/output/openapi.yaml",
        "path/to/output/ai-plugin.json",
        ["GET /api/resource"],
        AdaptiveCardUpdateStrategy.KeepExisting,
        undefined,
        false
      );

      assert.equal(fsCopyFileStub.callCount, 2);
      assert.isTrue(
        fsCopyFileStub.calledWith(
          pathMatcher("c:/tmp/working-dir/plugin/openapi.yaml"),
          pathMatcher("path/to/output/openapi.yaml")
        )
      );
      assert.isTrue(
        fsCopyFileStub.calledWith(
          pathMatcher("c:/tmp/working-dir/.kiota/documents/testapp/openapi.json"),
          pathMatcher("path/to/output/openapi.yaml.original")
        )
      );

      sinon.resetHistory();
      const result = await daSpecParser.generatePlugin(
        "path/to/spec.yaml",
        "path/to/manifest.json",
        "path/to/output/openapi.yaml",
        "path/to/output/ai-plugin.json",
        ["GET /api/resource"],
        AdaptiveCardUpdateStrategy.KeepExisting,
        undefined,
        true
      );

      assert.equal(fsCopyFileStub.callCount, 1);
      assert.isTrue(
        fsCopyFileStub.calledWith(
          pathMatcher("c:/tmp/working-dir/plugin/openapi.yaml"),
          pathMatcher("path/to/output/openapi.yaml")
        )
      );
    });

    it("should properly filter and merge functions when updating existing plugin", async () => {
      const readdirStub = sinon.stub(require("fs-extra"), "readdir").resolves(["openapi.json"]);
      const fsReadJSONStub = sinon.stub(require("fs-extra"), "readJSON");
      const fsPathExistsStub = sinon.stub(require("fs-extra"), "pathExists").resolves(true);
      const fsCopyStub = sinon.stub(require("fs-extra"), "copy").resolves();

      fsReadJSONStub.callsFake(async (path: any) => {
        if (path.includes("manifest.json")) {
          return { name: { short: "test-app" } };
        } else if (path.includes("ai-plugin.json") && path.includes("tmp")) {
          return {
            name: "test-app",
            runtimes: [
              {
                spec: { url: "path-to-be-replaced" },
                run_for_functions: ["newFunction1", "newFunction2"],
              },
            ],
            functions: [
              { name: "newFunction1", description: "New function 1" },
              { name: "newFunction2", description: "New function 2" },
            ],
          };
        } else {
          return {
            name: "test-app",
            runtimes: [
              {
                spec: { url: "../openapi.yaml" },
                run_for_functions: ["oldFunction1", "oldFunction2"],
              },
              {
                spec: { url: "other-spec.yaml" },
                run_for_functions: ["keepFunction1"],
              },
            ],
            functions: [
              { name: "oldFunction1", description: "Old function 1" },
              { name: "oldFunction2", description: "Old function 2" },
              { name: "keepFunction1", description: "Keep function 1" },
            ],
          };
        }
      });

      const fsWriteJsonStub = sinon.stub(require("fs-extra"), "writeJson").resolves();

      pathRelativeStub.returns("../openapi.yaml");

      const specPath = "path/to/spec.yaml";
      const teamsManifestPath = "path/to/manifest.json";
      const outputAPISpecPath = "path/to/output/openapi.yaml";
      const outputAIPluginPath = "path/to/output/ai-plugin.json";
      const operations = ["GET /api/resource"];

      const result = await daSpecParser.generatePlugin(
        specPath,
        teamsManifestPath,
        outputAPISpecPath,
        outputAIPluginPath,
        operations,
        AdaptiveCardUpdateStrategy.KeepExisting,
        undefined,
        true
      );

      assert.isTrue(
        fsWriteJsonStub.calledWith(
          pathMatcher("path/to/output/ai-plugin.json"),
          sinon.match((value) => {
            const hasNoOldFunctions = value.functions.every(
              (f: any) => f.name !== "oldFunction1" && f.name !== "oldFunction2"
            );

            const hasNewFunctions =
              value.functions.some((f: any) => f.name === "newFunction1") &&
              value.functions.some((f: any) => f.name === "newFunction2");

            const preservedOtherFunctions = value.functions.some(
              (f: any) => f.name === "keepFunction1"
            );

            const correctFunctionCount = value.functions.length === 3;

            const preservedOtherRuntime = value.runtimes.some(
              (r: any) =>
                r.spec.url === "other-spec.yaml" && r.run_for_functions.includes("keepFunction1")
            );

            const addedNewRuntime = value.runtimes.some(
              (r: any) =>
                r.spec.url === "../openapi.yaml" &&
                r.run_for_functions.includes("newFunction1") &&
                r.run_for_functions.includes("newFunction2")
            );

            return (
              hasNoOldFunctions &&
              hasNewFunctions &&
              preservedOtherFunctions &&
              correctFunctionCount &&
              preservedOtherRuntime &&
              addedNewRuntime
            );
          }),
          { spaces: 4 }
        )
      );

      assert.equal(fsCopyStub.callCount, 2);
      const copyCallArgs = fsCopyStub.secondCall.args;
      assert.isTrue(copyCallArgs[0].replace(/\\/g, "/").endsWith("adaptiveCards"));
      assert.isTrue(copyCallArgs[0].replace(/\\/g, "/").endsWith("adaptiveCards"));
      assert.deepEqual(
        copyCallArgs[2],
        {
          overwrite: false,
          errorOnExist: false,
        },
        "Copy options don't match"
      );

      assert.isFalse(
        fsCopyStub.firstCall.calledWith(
          pathMatcher("c:/tmp/working-dir/.kiota/documents/testapp/openapi.json"),
          sinon.match.any
        )
      );
    });
  });

  describe("patchOpenApiExtensionsIntoPluginManifest (issue #15731)", () => {
    let tmpDir: string;

    beforeEach(async () => {
      // The outer describe sets up sinon stubs we don't need here; restore so
      // real fs / yaml are exercised.
      sinon.restore();
      tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "kiota-15731-test-"));
    });

    afterEach(async () => {
      try {
        await fs.remove(tmpDir);
      } catch {
        // Ignore cleanup errors on Windows when files are still locked.
      }
    });

    async function writeSpec(yaml: string): Promise<string> {
      const specPath = path.join(tmpDir, "openapi.yaml");
      await fs.writeFile(specPath, yaml, "utf8");
      return specPath;
    }

    async function writeManifest(manifest: any): Promise<string> {
      const p = path.join(tmpDir, "test-apiplugin.json");
      await fs.writeJson(p, manifest, { spaces: 2 });
      return p;
    }

    it("propagates x-ai-adaptive-card and inlines the static_template", async () => {
      const cardsDir = path.join(tmpDir, "adaptiveCards");
      await fs.ensureDir(cardsDir);
      const card = { type: "AdaptiveCard", version: "1.5", body: [] };
      await fs.writeJson(path.join(cardsDir, "get.json"), card);

      const specPath = await writeSpec(
        [
          "openapi: 3.0.0",
          "info: { title: t, version: '1' }",
          "paths:",
          "  /items:",
          "    get:",
          "      operationId: getItems",
          "      x-ai-adaptive-card:",
          "        data_path: $.value[0]",
          "        file: adaptiveCards/get.json",
          "      responses: { '200': { description: ok } }",
          "",
        ].join("\n")
      );
      const manifestPath = await writeManifest({
        schema_version: "v2.4",
        functions: [{ name: "getItems", description: "" }],
      });

      await daSpecParser.patchOpenApiExtensionsIntoPluginManifest(specPath, manifestPath);

      const result = await fs.readJson(manifestPath);
      const fn = result.functions[0];
      assert.deepEqual(fn.capabilities.response_semantics.data_path, "$.value[0]");
      assert.deepEqual(fn.capabilities.response_semantics.static_template, card);
    });

    it("propagates x-openai-isConsequential into confirmation.isNonConsequential", async () => {
      const specPath = await writeSpec(
        [
          "openapi: 3.0.0",
          "info: { title: t, version: '1' }",
          "paths:",
          "  /items:",
          "    patch:",
          "      operationId: updateItems",
          "      x-openai-isConsequential: true",
          "      x-ai-capabilities:",
          "        confirmation:",
          "          type: AdaptiveCard",
          "          title: Update?",
          "          body: Confirm.",
          "      responses: { '200': { description: ok } }",
          "",
        ].join("\n")
      );
      const manifestPath = await writeManifest({
        schema_version: "v2.4",
        functions: [{ name: "updateItems", description: "" }],
      });

      await daSpecParser.patchOpenApiExtensionsIntoPluginManifest(specPath, manifestPath);

      const fn = (await fs.readJson(manifestPath)).functions[0];
      assert.equal(fn.capabilities.confirmation.type, "AdaptiveCard");
      assert.equal(fn.capabilities.confirmation.title, "Update?");
      // x-openai-isConsequential: true => isNonConsequential: false
      assert.strictEqual(fn.capabilities.confirmation.isNonConsequential, false);
    });

    it("maps x-openai-isConsequential: false to isNonConsequential: true", async () => {
      const specPath = await writeSpec(
        [
          "openapi: 3.0.0",
          "info: { title: t, version: '1' }",
          "paths:",
          "  /items:",
          "    post:",
          "      operationId: readItems",
          "      x-openai-isConsequential: false",
          "      x-ai-capabilities:",
          "        confirmation:",
          "          type: AdaptiveCard",
          "          title: Read",
          "          body: ok",
          "      responses: { '200': { description: ok } }",
          "",
        ].join("\n")
      );
      const manifestPath = await writeManifest({
        schema_version: "v2.4",
        functions: [{ name: "readItems", description: "" }],
      });

      await daSpecParser.patchOpenApiExtensionsIntoPluginManifest(specPath, manifestPath);

      const fn = (await fs.readJson(manifestPath)).functions[0];
      assert.strictEqual(fn.capabilities.confirmation.isNonConsequential, true);
    });

    it("does not overwrite capabilities Kiota already populated", async () => {
      const specPath = await writeSpec(
        [
          "openapi: 3.0.0",
          "info: { title: t, version: '1' }",
          "paths:",
          "  /items:",
          "    get:",
          "      operationId: getItems",
          "      x-ai-adaptive-card:",
          "        data_path: $.value[0]",
          "        file: adaptiveCards/get.json",
          "      responses: { '200': { description: ok } }",
          "",
        ].join("\n")
      );
      const existing = {
        schema_version: "v2.4",
        functions: [
          {
            name: "getItems",
            description: "",
            capabilities: {
              response_semantics: { data_path: "$.preserved" },
            },
          },
        ],
      };
      const manifestPath = await writeManifest(existing);

      await daSpecParser.patchOpenApiExtensionsIntoPluginManifest(specPath, manifestPath);

      const fn = (await fs.readJson(manifestPath)).functions[0];
      assert.equal(fn.capabilities.response_semantics.data_path, "$.preserved");
    });

    it("is a no-op when the spec has no relevant extensions", async () => {
      const specPath = await writeSpec(
        [
          "openapi: 3.0.0",
          "info: { title: t, version: '1' }",
          "paths:",
          "  /items:",
          "    get:",
          "      operationId: getItems",
          "      responses: { '200': { description: ok } }",
          "",
        ].join("\n")
      );
      const before = {
        schema_version: "v2.4",
        functions: [{ name: "getItems", description: "" }],
      };
      const manifestPath = await writeManifest(before);

      await daSpecParser.patchOpenApiExtensionsIntoPluginManifest(specPath, manifestPath);

      assert.deepEqual(await fs.readJson(manifestPath), before);
    });

    it("returns silently when spec or manifest path does not exist", async () => {
      // Should not throw.
      await daSpecParser.patchOpenApiExtensionsIntoPluginManifest(
        path.join(tmpDir, "missing.yaml"),
        path.join(tmpDir, "missing.json")
      );
    });

    it("returns silently when the spec is not valid YAML", async () => {
      const specPath = path.join(tmpDir, "openapi.yaml");
      // `: : :` is a YAML parse error.
      await fs.writeFile(specPath, ":\n: :\n  : :", "utf8");
      const before = {
        schema_version: "v2.4",
        functions: [{ name: "x", description: "" }],
      };
      const manifestPath = await writeManifest(before);

      await daSpecParser.patchOpenApiExtensionsIntoPluginManifest(specPath, manifestPath);

      assert.deepEqual(await fs.readJson(manifestPath), before);
    });

    it("returns silently when the parsed spec has no paths object", async () => {
      // YAML parses to a plain string, which exercises the
      // `!spec || typeof spec !== "object" || !spec.paths` guard.
      const specPath = await writeSpec("just-a-string\n");
      const before = {
        schema_version: "v2.4",
        functions: [{ name: "x", description: "" }],
      };
      const manifestPath = await writeManifest(before);

      await daSpecParser.patchOpenApiExtensionsIntoPluginManifest(specPath, manifestPath);

      assert.deepEqual(await fs.readJson(manifestPath), before);
    });

    it("leaves static_template unset when the referenced card file is unreadable", async () => {
      // Create a directory where a card file is expected; readJson will fail.
      const cardsDir = path.join(tmpDir, "adaptiveCards");
      await fs.ensureDir(cardsDir);
      await fs.ensureDir(path.join(cardsDir, "broken.json"));

      const specPath = await writeSpec(
        [
          "openapi: 3.0.0",
          "info: { title: t, version: '1' }",
          "paths:",
          "  /items:",
          "    get:",
          "      operationId: getItems",
          "      x-ai-adaptive-card:",
          "        data_path: $.items[0]",
          "        file: adaptiveCards/broken.json",
          "      responses: { '200': { description: ok } }",
          "",
        ].join("\n")
      );
      const manifestPath = await writeManifest({
        schema_version: "v2.4",
        functions: [{ name: "getItems", description: "" }],
      });

      await daSpecParser.patchOpenApiExtensionsIntoPluginManifest(specPath, manifestPath);

      const fn = (await fs.readJson(manifestPath)).functions[0];
      // data_path is still propagated; static_template silently omitted.
      assert.equal(fn.capabilities.response_semantics.data_path, "$.items[0]");
      assert.notProperty(fn.capabilities.response_semantics, "static_template");
    });

    it("inlines Kiota's `{ file }` placeholder for static_template", async () => {
      // Kiota 1.31.1 emits `static_template: { file: "adaptiveCards/<name>.json" }`
      // instead of inlining the card. The patch must replace it with the
      // actual card JSON.
      const cardsDir = path.join(tmpDir, "adaptiveCards");
      await fs.ensureDir(cardsDir);
      const cardJson = {
        type: "AdaptiveCard",
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        version: "1.5",
        body: [{ type: "TextBlock", text: "${title}" }],
      };
      await fs.writeJson(path.join(cardsDir, "search.json"), cardJson);

      const specPath = await writeSpec(
        [
          "openapi: 3.0.0",
          "info: { title: t, version: '1' }",
          "paths:",
          "  /search:",
          "    get:",
          "      operationId: searchIssues",
          "      x-ai-adaptive-card:",
          "        data_path: $.items",
          "        file: adaptiveCards/search.json",
          "      responses: { '200': { description: ok } }",
          "",
        ].join("\n")
      );
      const manifestPath = await writeManifest({
        schema_version: "v2.4",
        functions: [
          {
            name: "searchIssues",
            description: "",
            capabilities: {
              response_semantics: {
                data_path: "$.items",
                static_template: { file: "adaptiveCards/search.json" },
              },
            },
          },
        ],
      });

      await daSpecParser.patchOpenApiExtensionsIntoPluginManifest(specPath, manifestPath);

      const fn = (await fs.readJson(manifestPath)).functions[0];
      assert.deepEqual(fn.capabilities.response_semantics.static_template, cardJson);
      assert.equal(fn.capabilities.response_semantics.data_path, "$.items");
    });

    it("does not replace a real static_template that already has card fields", async () => {
      // If `static_template` already looks like a real Adaptive Card (has
      // `type`/`$schema`/`body`), we must leave it alone even if a `file`
      // property happens to be present.
      const cardsDir = path.join(tmpDir, "adaptiveCards");
      await fs.ensureDir(cardsDir);
      await fs.writeJson(path.join(cardsDir, "search.json"), {
        type: "AdaptiveCard",
        body: [{ type: "TextBlock", text: "from disk" }],
      });

      const specPath = await writeSpec(
        [
          "openapi: 3.0.0",
          "info: { title: t, version: '1' }",
          "paths:",
          "  /search:",
          "    get:",
          "      operationId: searchIssues",
          "      x-ai-adaptive-card:",
          "        data_path: $.items",
          "        file: adaptiveCards/search.json",
          "      responses: { '200': { description: ok } }",
          "",
        ].join("\n")
      );
      const realCard = {
        type: "AdaptiveCard",
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        body: [{ type: "TextBlock", text: "preserved" }],
      };
      const manifestPath = await writeManifest({
        schema_version: "v2.4",
        functions: [
          {
            name: "searchIssues",
            description: "",
            capabilities: {
              response_semantics: {
                data_path: "$.items",
                static_template: realCard,
              },
            },
          },
        ],
      });

      await daSpecParser.patchOpenApiExtensionsIntoPluginManifest(specPath, manifestPath);

      const fn = (await fs.readJson(manifestPath)).functions[0];
      assert.deepEqual(fn.capabilities.response_semantics.static_template, realCard);
    });

    it("resolves card files from the parent of the plugin manifest directory", async () => {
      // Mirrors the real layout: spec lives in `<root>/.generated/specs/`,
      // plugin manifest lives in `<root>/.generated/`, and the card lives in
      // `<root>/adaptiveCards/`.
      const generated = path.join(tmpDir, ".generated");
      const specsDir = path.join(generated, "specs");
      await fs.ensureDir(specsDir);
      await fs.ensureDir(path.join(tmpDir, "adaptiveCards"));
      const cardJson = { type: "AdaptiveCard", body: [] };
      await fs.writeJson(path.join(tmpDir, "adaptiveCards", "card.json"), cardJson);

      const specPath = path.join(specsDir, "openapi.yaml");
      await fs.writeFile(
        specPath,
        [
          "openapi: 3.0.0",
          "info: { title: t, version: '1' }",
          "paths:",
          "  /x:",
          "    get:",
          "      operationId: getX",
          "      x-ai-adaptive-card:",
          "        data_path: $.items",
          "        file: adaptiveCards/card.json",
          "      responses: { '200': { description: ok } }",
          "",
        ].join("\n"),
        "utf8"
      );
      const manifestPath = path.join(generated, "x-apiplugin.json");
      await fs.writeJson(manifestPath, {
        schema_version: "v2.4",
        functions: [
          {
            name: "getX",
            description: "",
            capabilities: {
              response_semantics: {
                data_path: "$.items",
                static_template: { file: "adaptiveCards/card.json" },
              },
            },
          },
        ],
      });

      await daSpecParser.patchOpenApiExtensionsIntoPluginManifest(specPath, manifestPath);

      const fn = (await fs.readJson(manifestPath)).functions[0];
      assert.deepEqual(fn.capabilities.response_semantics.static_template, cardJson);
    });

    it("swallows readJSON errors when the card file exists but is invalid JSON", async () => {
      // Exercises the `catch` branch inside resolveCardJson: pathExists
      // succeeds but readJSON throws on a corrupt file.
      const cardsDir = path.join(tmpDir, "adaptiveCards");
      await fs.ensureDir(cardsDir);
      await fs.writeFile(path.join(cardsDir, "broken.json"), "{ not valid json", "utf8");

      const specPath = await writeSpec(
        [
          "openapi: 3.0.0",
          "info: { title: t, version: '1' }",
          "paths:",
          "  /items:",
          "    get:",
          "      operationId: getItems",
          "      x-ai-adaptive-card:",
          "        data_path: $.items",
          "        file: adaptiveCards/broken.json",
          "      responses: { '200': { description: ok } }",
          "",
        ].join("\n")
      );
      const manifestPath = await writeManifest({
        schema_version: "v2.4",
        functions: [
          {
            name: "getItems",
            description: "",
            capabilities: {
              response_semantics: {
                static_template: { file: "adaptiveCards/broken.json" },
              },
            },
          },
        ],
      });

      await daSpecParser.patchOpenApiExtensionsIntoPluginManifest(specPath, manifestPath);

      // The placeholder should be left intact since the card could not be
      // read; data_path should still get filled from the spec.
      const fn = (await fs.readJson(manifestPath)).functions[0];
      assert.deepEqual(fn.capabilities.response_semantics.static_template, {
        file: "adaptiveCards/broken.json",
      });
      assert.equal(fn.capabilities.response_semantics.data_path, "$.items");
    });

    it("fills missing data_path on an existing response_semantics with a placeholder", async () => {
      // response_semantics exists with the Kiota `{ file }` placeholder but
      // without a data_path; the patcher should backfill data_path from the
      // spec while also inlining the card.
      const cardsDir = path.join(tmpDir, "adaptiveCards");
      await fs.ensureDir(cardsDir);
      const card = { type: "AdaptiveCard", version: "1.5", body: [] };
      await fs.writeJson(path.join(cardsDir, "get.json"), card);

      const specPath = await writeSpec(
        [
          "openapi: 3.0.0",
          "info: { title: t, version: '1' }",
          "paths:",
          "  /items:",
          "    get:",
          "      operationId: getItems",
          "      x-ai-adaptive-card:",
          "        data_path: $.value",
          "        file: adaptiveCards/get.json",
          "      responses: { '200': { description: ok } }",
          "",
        ].join("\n")
      );
      const manifestPath = await writeManifest({
        schema_version: "v2.4",
        functions: [
          {
            name: "getItems",
            description: "",
            capabilities: {
              response_semantics: {
                // No data_path; placeholder static_template.
                static_template: { file: "adaptiveCards/get.json" },
              },
            },
          },
        ],
      });

      await daSpecParser.patchOpenApiExtensionsIntoPluginManifest(specPath, manifestPath);

      const fn = (await fs.readJson(manifestPath)).functions[0];
      assert.equal(fn.capabilities.response_semantics.data_path, "$.value");
      assert.deepEqual(fn.capabilities.response_semantics.static_template, card);
    });

    it("fills missing data_path on an existing response_semantics with a real card", async () => {
      // response_semantics exists with a real (non-placeholder) static_template
      // but no data_path. The patcher should leave the card alone and only
      // backfill data_path.
      const realCard = {
        type: "AdaptiveCard",
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        body: [{ type: "TextBlock", text: "real" }],
      };
      const specPath = await writeSpec(
        [
          "openapi: 3.0.0",
          "info: { title: t, version: '1' }",
          "paths:",
          "  /items:",
          "    get:",
          "      operationId: getItems",
          "      x-ai-adaptive-card:",
          "        data_path: $.value",
          "        file: adaptiveCards/get.json",
          "      responses: { '200': { description: ok } }",
          "",
        ].join("\n")
      );
      const manifestPath = await writeManifest({
        schema_version: "v2.4",
        functions: [
          {
            name: "getItems",
            description: "",
            capabilities: {
              response_semantics: {
                static_template: realCard,
              },
            },
          },
        ],
      });

      await daSpecParser.patchOpenApiExtensionsIntoPluginManifest(specPath, manifestPath);

      const fn = (await fs.readJson(manifestPath)).functions[0];
      assert.equal(fn.capabilities.response_semantics.data_path, "$.value");
      assert.deepEqual(fn.capabilities.response_semantics.static_template, realCard);
    });
  });
});

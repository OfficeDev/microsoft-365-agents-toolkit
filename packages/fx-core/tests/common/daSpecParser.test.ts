// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import "mocha";
import sinon from "sinon";
import * as kiotaClient from "../../src/common/kiotaClient";
import * as daSpecParser from "../../src/common/daSpecParser";
import * as utils from "../../src/common/utils";
import { featureFlagManager, FeatureFlags } from "../../src/common/featureFlags";
import { Platform } from "@microsoft/teamsfx-api";
import { KiotaOpenApiNode, KiotaTreeResult } from "@microsoft/kiota";
import {
  AdaptiveCardUpdateStrategy,
  ErrorType,
  ValidationStatus,
  WarningType,
} from "@microsoft/m365-spec-parser";

describe("daSpecParser", () => {
  let listAPITreeInfoStub: sinon.SinonStub;
  let featureFlagStub: sinon.SinonStub;
  let isJsonSpecFileStub: sinon.SinonStub;

  beforeEach(() => {
    listAPITreeInfoStub = sinon.stub(kiotaClient, "listAPITreeInfo");
    featureFlagStub = sinon.stub(featureFlagManager, "getBooleanValue");
    isJsonSpecFileStub = sinon.stub(utils, "isJsonSpecFile");
    featureFlagStub.withArgs(FeatureFlags.KiotaNPMIntegration).returns(true);
    isJsonSpecFileStub.resolves(false);
  });

  afterEach(() => {
    sinon.restore();
  });

  describe("listAPIInfo with KiotaNPMIntegration enabled", () => {
    it("should return empty result when treeInfo is undefined", async () => {
      listAPITreeInfoStub.resolves(undefined);

      const result = await daSpecParser.listAPIInfo("path/to/spec");

      assert.deepEqual(result, {
        allAPICount: 0,
        validAPICount: 0,
        APIs: [],
      });
    });

    it("should return empty result when rootNode is undefined", async () => {
      listAPITreeInfoStub.resolves({ rootNode: undefined });

      const result = await daSpecParser.listAPIInfo("path/to/spec");

      assert.deepEqual(result, {
        allAPICount: 0,
        validAPICount: 0,
        APIs: [],
      });
    });

    it("should extract operations with simple node structure", async () => {
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

    it("should handle Windows-style paths with backslashes", async () => {
      const mockTreeInfo: KiotaTreeResult = {
        rootNode: {
          isOperation: true,
          path: "api\\resource",
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
              path: "api/users",
              segment: "GET",
              operationId: "getUsers",
              summary: "Get users",
              description: "Get all users",
              selected: true,
              children: [],
            },
            {
              isOperation: true,
              path: "api/posts",
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
                  path: "api/users/profile",
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
      };

      listAPITreeInfoStub.resolves(mockTreeInfo);

      const result = await daSpecParser.listAPIInfo("path/to/spec");

      assert.equal(result.APIs[0].auth?.name, "api_key, oauth2");
      assert.deepEqual(result.APIs[0].auth?.authScheme, {
        type: "multipleAuth",
      });
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
        { type: ErrorType.SpecNotValid, content: `Error: ${errorMessage}` },
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
              path: "api/resource",
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
      assert.isEmpty(result.warnings);

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
    });
  });

  describe("generatePlugin with KiotaNPMIntegration enabled", () => {
    let kiotaGeneratePluginStub: sinon.SinonStub;
    let tmpDirSyncStub: sinon.SinonStub;
    let fsReadJSONStub: sinon.SinonStub;
    let fsCopyFileStub: sinon.SinonStub;
    let fsWriteJsonStub: sinon.SinonStub;
    let parseAndUpdatePluginManifestStub: sinon.SinonStub;
    let pathRelativeStub: sinon.SinonStub;

    beforeEach(() => {
      kiotaGeneratePluginStub = sinon.stub(kiotaClient, "kiotageneratePlugin");
      tmpDirSyncStub = sinon.stub(require("tmp"), "dirSync");
      fsReadJSONStub = sinon.stub(require("fs-extra"), "readJSON");
      fsCopyFileStub = sinon.stub(require("fs-extra"), "copyFile");
      fsWriteJsonStub = sinon.stub(require("fs-extra"), "writeJson");
      parseAndUpdatePluginManifestStub = sinon.stub(
        daSpecParser,
        "parseAndUpdatePluginManifestForKiota"
      );
      pathRelativeStub = sinon.stub(require("path"), "relative");

      featureFlagStub.withArgs(FeatureFlags.KiotaNPMIntegration).returns(true);

      // Setup common stubs
      tmpDirSyncStub.returns({
        name: "c:\\tmp\\working-dir",
        removeCallback: sinon.stub(),
        unsafeCleanup: true,
      });
      fsReadJSONStub.resolves({ name: { short: "test-app" } });
      kiotaGeneratePluginStub.resolves({
        openAPISpec: "c:\\tmp\\working-dir\\plugin\\openapi.yaml",
        aiPlugin: "c:\\tmp\\working-dir\\plugin\\ai-plugin.json",
        logs: [],
      });
      fsCopyFileStub.resolves();
      fsWriteJsonStub.resolves();
      parseAndUpdatePluginManifestStub.resolves([]);
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

      assert.equal(fsCopyFileStub.callCount, 3);

      assert.isTrue(
        fsCopyFileStub.firstCall.calledWith(
          pathMatcher("c:/tmp/working-dir/plugin/openapi.yaml"),
          pathMatcher("path/to/output/openapi.yaml")
        )
      );
      assert.isTrue(
        fsCopyFileStub.secondCall.calledWith(
          pathMatcher("c:/tmp/working-dir/plugin/ai-plugin.json"),
          pathMatcher("path/to/output/ai-plugin.json")
        )
      );

      assert.isTrue(
        fsCopyFileStub.thirdCall.calledWith(
          pathMatcher("path/to/spec.yaml"),
          pathMatcher("path/to/output/openapi.original.yaml")
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

      isJsonSpecFileStub.withArgs(specPath).resolves(true);

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
      assert.isTrue(result.warnings.some((w) => w.type === WarningType.OperationIdMissing));
      assert.isTrue(
        result.warnings.some((w) => w.type === WarningType.OperationIdContainsSpecialCharacters)
      );
      assert.isTrue(result.warnings.some((w) => w.type === WarningType.UnsupportedAuthType));

      assert.isTrue(
        fsCopyFileStub.calledWith(
          pathMatcher("c:/tmp/working-dir/plugin/openapi.yaml"),
          pathMatcher("path/to/output/openapi.yaml")
        )
      );

      assert.isTrue(
        fsCopyFileStub.calledWith(
          pathMatcher("path/to/spec.json"),
          pathMatcher("path/to/output/openapi.original.json")
        )
      );
    });

    it("should handle manifest with environment variables and special characters", async () => {
      const complexManifest = {
        name: {
          short: "Complex$App-Name_${{ENV_VAR}}",
        },
      };
      fsReadJSONStub.resolves(complexManifest);

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
        ["GET /api"],
        AdaptiveCardUpdateStrategy.KeepExisting
      );

      const expectedPluginManifest = {
        name: { short: "test-app" },
        runtimes: [
          {
            spec: { url: "../../openapi.yaml" },
          },
        ],
      };

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
      fsReadJSONStub.resolves(pluginManifest);

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
      isJsonSpecFileStub.withArgs("path/to/spec.json").resolves(true);

      await daSpecParser.generatePlugin(
        "path/to/spec.json",
        "path/to/manifest.json",
        "path/to/output/openapi.spec",
        "path/to/output/ai-plugin.json",
        ["GET /api"],
        AdaptiveCardUpdateStrategy.KeepExisting
      );

      assert.isTrue(
        fsCopyFileStub.calledWith(
          pathMatcher("path/to/spec.json"),
          pathMatcher("path/to/output/openapi.original.json")
        )
      );

      sinon.resetHistory();

      isJsonSpecFileStub.withArgs("path/to/spec.yaml").resolves(false);

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
          pathMatcher("path/to/spec.yaml"),
          pathMatcher("path/to/output/openapi.original.yaml")
        )
      );
    });
  });
});

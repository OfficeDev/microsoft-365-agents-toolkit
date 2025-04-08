// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import "mocha";
import sinon from "sinon";
import * as kiotaClient from "../../src/common/kiotaClient";
import * as daSpecParser from "../../src/common/daSpecParser";
import { featureFlagManager, FeatureFlags } from "../../src/common/featureFlags";
import { Platform } from "@microsoft/teamsfx-api";
import { KiotaOpenApiNode, KiotaTreeResult } from "@microsoft/kiota";

describe("daSpecParser", () => {
  let listAPITreeInfoStub: sinon.SinonStub;
  let featureFlagStub: sinon.SinonStub;

  beforeEach(() => {
    listAPITreeInfoStub = sinon.stub(kiotaClient, "listAPITreeInfo");
    featureFlagStub = sinon.stub(featureFlagManager, "getBooleanValue");
    featureFlagStub.withArgs(FeatureFlags.KiotaNPMIntegration).returns(true);
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
              children: [],
            },
            {
              isOperation: true,
              path: "api/posts",
              segment: "POST",
              operationId: "createPost",
              summary: "Create post",
              description: "Create a new post",
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
  });
});

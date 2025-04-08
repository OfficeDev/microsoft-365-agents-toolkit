// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import "mocha";
import sinon from "sinon";
import { searchOpenAPISpec } from "../../src/common/kiotaClient";
import proxyquire from "proxyquire";

describe("kiotaClient", () => {
  const sandbox = sinon.createSandbox();
  afterEach(async () => {
    sandbox.restore();
  });

  it("happy path: searchOpenAPISpec", async () => {
    const mockSearchResult = {
      "api-spec": {
        DescriptionUrl: "https://example.com/api-spec.json",
        Description: "API Spec description",
        Title: "API Spec Title",
      },
    };

    process.env.KIOTA_BINARY_PATH = "mock/path/to/kiota";

    const setKiotaConfigStub = sinon.stub().resolves();
    const searchDescriptionStub = sinon.stub().resolves(mockSearchResult);

    const { searchOpenAPISpec } = proxyquire("../../src/common/kiotaClient", {
      "@microsoft/kiota": {
        setKiotaConfig: setKiotaConfigStub,
        searchDescription: searchDescriptionStub,
        "@noCallThru": true,
      },
    });

    const result = await searchOpenAPISpec("test-query");

    assert(setKiotaConfigStub.calledOnce);
    assert(setKiotaConfigStub.calledWith({ binaryLocation: "mock/path/to/kiota" }));
    assert(searchDescriptionStub.calledOnce);
    assert(
      searchDescriptionStub.calledWith({
        searchTerm: "test-query",
        clearCache: false,
      })
    );

    assert.equal(result.length, 1);
    assert.equal(result[0].key, "api-spec");
    assert.equal(result[0].url, "https://example.com/api-spec.json");
    assert.equal(result[0].description, "API Spec description");
  });

  it("happy path: searchOpenAPISpec missing url", async () => {
    const mockSearchResult = {
      Description: "API Spec description",
      Title: "API Spec Title",
    };

    if (process.env.KIOTA_BINARY_PATH) {
      delete process.env.KIOTA_BINARY_PATH;
    }

    const setKiotaConfigStub = sinon.stub().resolves();
    const searchDescriptionStub = sinon.stub().resolves(mockSearchResult);

    const { searchOpenAPISpec } = proxyquire("../../src/common/kiotaClient", {
      "@microsoft/kiota": {
        setKiotaConfig: setKiotaConfigStub,
        searchDescription: searchDescriptionStub,
        "@noCallThru": true,
      },
    });

    const result = await searchOpenAPISpec("test-query");
    assert(setKiotaConfigStub.notCalled);
    assert.equal(result.length, 0);
  });

  it("happy path: searchOpenAPISpec undefined result", async () => {
    if (process.env.KIOTA_BINARY_PATH) {
      delete process.env.KIOTA_BINARY_PATH;
    }

    const setKiotaConfigStub = sinon.stub().resolves();
    const searchDescriptionStub = sinon.stub().resolves(undefined);

    const { searchOpenAPISpec } = proxyquire("../../src/common/kiotaClient", {
      "@microsoft/kiota": {
        setKiotaConfig: setKiotaConfigStub,
        searchDescription: searchDescriptionStub,
        "@noCallThru": true,
      },
    });

    const result = await searchOpenAPISpec("test-query");
    assert(setKiotaConfigStub.notCalled);
    assert.equal(result.length, 0);
  });

  describe("listAPITreeInfo", () => {
    const sandbox = sinon.createSandbox();
    afterEach(async () => {
      sandbox.restore();
    });

    it("happy path: listAPITreeInfo with default parameters", async () => {
      const mockTreeResult = {
        rootNode: {
          isOperation: true,
          path: "api/resource",
          segment: "GET",
          operationId: "getResource",
          children: [],
        },
        servers: ["https://api.example.com"],
        security: [],
        securitySchemes: {},
        logs: [],
      };

      process.env.KIOTA_BINARY_PATH = "mock/path/to/kiota";

      const setKiotaConfigStub = sinon.stub().resolves();
      const getKiotaTreeStub = sinon.stub().resolves(mockTreeResult);

      const { listAPITreeInfo } = proxyquire("../../src/common/kiotaClient", {
        "@microsoft/kiota": {
          getKiotaTree: getKiotaTreeStub,
          setKiotaConfig: setKiotaConfigStub,
          "@noCallThru": true,
        },
      });

      const result = await listAPITreeInfo("path/to/spec");

      assert(getKiotaTreeStub.calledOnce);
      assert(
        getKiotaTreeStub.calledWith({
          includeFilters: undefined,
          descriptionPath: "path/to/spec",
          excludeFilters: undefined,
          clearCache: true,
          includeKiotaValidationRules: true,
        })
      );

      assert.deepEqual(result, mockTreeResult);
    });

    it("happy path: listAPITreeInfo with include and exclude filters", async () => {
      const mockTreeResult = {
        rootNode: {
          isOperation: false,
          path: "",
          segment: "",
          children: [
            {
              isOperation: true,
              path: "api/users",
              segment: "GET",
              operationId: "getUsers",
              children: [],
            },
          ],
        },
        servers: ["https://api.example.com"],
        security: [],
        securitySchemes: {},
        logs: [],
      };

      const includeFilters = ["GET /users"];
      const excludeFilters = ["DELETE /users"];

      const getKiotaTreeStub = sinon.stub().resolves(mockTreeResult);
      const setKiotaConfigStub = sinon.stub().resolves();

      const { listAPITreeInfo } = proxyquire("../../src/common/kiotaClient", {
        "@microsoft/kiota": {
          getKiotaTree: getKiotaTreeStub,
          setKiotaConfig: setKiotaConfigStub,
          "@noCallThru": true,
        },
      });

      const result = await listAPITreeInfo("path/to/spec", includeFilters, excludeFilters);

      assert(getKiotaTreeStub.calledOnce);
      assert(
        getKiotaTreeStub.calledWith({
          includeFilters: includeFilters,
          descriptionPath: "path/to/spec",
          excludeFilters: excludeFilters,
          clearCache: true,
          includeKiotaValidationRules: true,
        })
      );

      assert.deepEqual(result, mockTreeResult);
    });

    it("edge case: listAPITreeInfo returns undefined", async () => {
      const getKiotaTreeStub = sinon.stub().resolves(undefined);

      const setKiotaConfigStub = sinon.stub().resolves();

      const { listAPITreeInfo } = proxyquire("../../src/common/kiotaClient", {
        "@microsoft/kiota": {
          getKiotaTree: getKiotaTreeStub,
          setKiotaConfig: setKiotaConfigStub,
          "@noCallThru": true,
        },
      });

      const result = await listAPITreeInfo("path/to/spec");

      assert(getKiotaTreeStub.calledOnce);
      assert.isUndefined(result);
    });

    it("error path: listAPITreeInfo throws exception", async () => {
      const errorMessage = "Failed to parse OpenAPI spec";
      const getKiotaTreeStub = sinon.stub().rejects(new Error(errorMessage));

      const setKiotaConfigStub = sinon.stub().resolves();

      const { listAPITreeInfo } = proxyquire("../../src/common/kiotaClient", {
        "@microsoft/kiota": {
          getKiotaTree: getKiotaTreeStub,
          setKiotaConfig: setKiotaConfigStub,
          "@noCallThru": true,
        },
      });

      try {
        await listAPITreeInfo("path/to/spec");
        assert.fail("Should have thrown an error");
      } catch (error) {
        assert.equal((error as Error).message, errorMessage);
      }
    });
  });
});

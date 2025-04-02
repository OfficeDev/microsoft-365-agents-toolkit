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
});

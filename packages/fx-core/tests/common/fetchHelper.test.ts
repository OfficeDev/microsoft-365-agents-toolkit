// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";
import { assert } from "chai";
import sinon from "sinon";
import * as fetchHelper from "../../src/common/fetchHelper";

describe("fetchHelper", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("should export a default function", () => {
    assert.isFunction(fetchHelper.default);
  });

  it("should be stubbable via sinon", async () => {
    const fakeResponse = { ok: true, status: 200, json: async () => ({ key: "value" }) };
    sandbox.stub(fetchHelper, "default").resolves(fakeResponse as any);

    const result = await fetchHelper.default("https://example.com");
    assert.strictEqual(result.ok, true);
    assert.strictEqual(result.status, 200);
    const json = await result.json();
    assert.deepStrictEqual(json, { key: "value" });
  });

  it("should pass url and init to the stubbed function", async () => {
    const stub = sandbox.stub(fetchHelper, "default").resolves({ ok: true } as any);

    const init = { method: "POST", headers: { "Content-Type": "application/json" } };
    await fetchHelper.default("https://example.com/api", init as any);

    assert.isTrue(stub.calledOnce);
    assert.strictEqual(stub.firstCall.args[0], "https://example.com/api");
    assert.deepStrictEqual(stub.firstCall.args[1], init);
  });

  it("should call real node-fetch when not stubbed", async () => {
    // Verify the real function returns a promise (doesn't throw synchronously)
    const promise = fetchHelper.default("https://localhost:0");
    assert.instanceOf(promise, Promise);
    // The fetch will fail (connection refused) but should reject, not throw
    try {
      await promise;
      assert.fail("Expected fetch to reject for invalid URL");
    } catch (err: any) {
      // Expected: connection refused or fetch error
      assert.isOk(err);
    }
  });
});

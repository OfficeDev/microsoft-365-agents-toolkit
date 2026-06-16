// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
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
    sandbox.stub(globalThis, "fetch").resolves(fakeResponse as any);

    const result = await fetchHelper.default("https://example.com");
    assert.strictEqual(result.ok, true);
    assert.strictEqual(result.status, 200);
    const json = await result.json();
    assert.deepStrictEqual(json, { key: "value" });
  });

  it("should pass url and init to the stubbed function", async () => {
    const stub = sandbox.stub(globalThis, "fetch").resolves({ ok: true } as any);

    const init = { method: "POST", headers: { "Content-Type": "application/json" } };
    await fetchHelper.default("https://example.com/api", init as any);

    assert.isTrue(stub.calledOnce);
    assert.strictEqual(stub.firstCall.args[0], "https://example.com/api");
    assert.deepStrictEqual(stub.firstCall.args[1], init);
  });

  it("should call real fetch when not stubbed", async () => {
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

  it("should use globalThis.fetch when available", async () => {
    const originalFetch = globalThis.fetch;
    try {
      const fakeResponse = { ok: true, status: 200 };
      globalThis.fetch = sandbox.stub().resolves(fakeResponse) as any;
      const result = await fetchHelper.default("https://example.com");
      assert.strictEqual(result.ok, true);
      assert.strictEqual(result.status, 200);
      assert.isTrue((globalThis.fetch as sinon.SinonStub).calledOnce);
      assert.strictEqual(
        (globalThis.fetch as sinon.SinonStub).firstCall.args[0],
        "https://example.com"
      );
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it("should fall back to node-fetch when globalThis.fetch is unavailable", async () => {
    const originalFetch = globalThis.fetch;
    try {
      // Remove globalThis.fetch to simulate environments where it is not available
      (globalThis as any).fetch = undefined;
      // Use an already-aborted signal to avoid long network waits in fallback path.
      const controller = new AbortController();
      controller.abort();
      const promise = fetchHelper.default("https://example.com", {
        signal: controller.signal,
      } as any);
      assert.instanceOf(promise, Promise);
      try {
        await promise;
      } catch (err: any) {
        // Expected: abort or fetch error from node-fetch fallback.
        assert.isOk(err);
        assert.notInclude(String(err?.message ?? ""), "fetch only supports HTTP/HTTPS URLs");
      }
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it("should reject non-HTTP(S) URLs", async () => {
    for (const badUrl of [
      "file:///etc/passwd",
      "ftp://example.com",
      "/local/path",
      "data:text/html,<h1>hi</h1>",
    ]) {
      try {
        await fetchHelper.default(badUrl);
        assert.fail(`Expected fetch to throw for URL: ${badUrl}`);
      } catch (err: any) {
        assert.include(err.message, "fetch only supports HTTP/HTTPS URLs");
      }
    }
  });

  it("should accept a URL object", async () => {
    const originalFetch = globalThis.fetch;
    try {
      const fakeResponse = { ok: true, status: 200 };
      globalThis.fetch = sandbox.stub().resolves(fakeResponse) as any;
      const urlObj = new URL("https://example.com/path");
      const result = await fetchHelper.default(urlObj);
      assert.strictEqual(result.ok, true);
      assert.isTrue((globalThis.fetch as sinon.SinonStub).calledOnce);
      assert.strictEqual(
        (globalThis.fetch as sinon.SinonStub).firstCall.args[0],
        "https://example.com/path"
      );
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

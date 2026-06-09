// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { expect } from "chai";
import { deterministicAppId } from "../../../../src/component/generator/openPlugin/deterministicId";

describe("openPlugin.deterministicAppId", () => {
  it("is deterministic for the same seed", () => {
    const a = deterministicAppId("openplugin:demo");
    const b = deterministicAppId("openplugin:demo");
    expect(a).to.equal(b);
  });

  it("differs for different seeds", () => {
    expect(deterministicAppId("openplugin:demo")).to.not.equal(
      deterministicAppId("openplugin:other")
    );
  });

  it("produces UUIDv5 (version 5 and RFC 4122 variant)", () => {
    const id = deterministicAppId("openplugin:demo");
    expect(id).to.match(/^[0-9a-f]{8}-[0-9a-f]{4}-5[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/);
  });

  it("matches the npm 'uuid' library's v5 output for the URL namespace + 'www.example.com'", () => {
    // Cross-checked against `uuid` npm package: v5("www.example.com", URL_NAMESPACE).
    // Pinned here as a regression anchor for the byte mangling and namespace bytes.
    expect(deterministicAppId("www.example.com")).to.equal("b63cdfa4-3df9-568e-97ae-006c5b8fd652");
  });
});

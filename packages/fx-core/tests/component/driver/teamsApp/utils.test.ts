// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { expect } from "chai";
import { normalizePath } from "../../../../src/component/driver/teamsApp/utils/utils";

describe("utils", async () => {
  it("normalizePath: should use forward slash", () => {
    const res = normalizePath("resources\\test.yaml", true);
    expect(res).equal("resources/test.yaml");
  });

  it("normalizePath: no need to convert", () => {
    const res = normalizePath("resources\\test.yaml", false);
    expect(res).equal("resources\\test.yaml");
  });
});

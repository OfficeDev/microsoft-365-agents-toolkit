// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { expect } from "chai";
import { isValidProject } from "../../src/common/projectSettingsHelper";
describe("tools", () => {
  it("is not valid project", () => {
    expect(isValidProject()).is.false;
  });
});

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as chai from "chai";
import { isValidHttpUrl } from "../../src/common/stringUtils";

describe("isValidHttpUrl", () => {
  it("valid https url", () => {
    const url = "https://www.bing.com";
    chai.expect(isValidHttpUrl(url)).equals(true);
  });

  it("valid http url", () => {
    const url = "http://www.bing.com";
    chai.expect(isValidHttpUrl(url)).equals(true);
  });

  it("invalid url", () => {
    const url = "abc";
    chai.expect(isValidHttpUrl(url)).equals(false);
  });

  it("invalid protocol", () => {
    const url = "vscode://www.bing.com";
    chai.expect(isValidHttpUrl(url)).equals(false);
  });
});

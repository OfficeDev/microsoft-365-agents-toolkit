// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import sinon from "sinon";
import { dictMatcher } from "../../src/common/secretmasker/dict";
import { secretMasker } from "../../src/common/secretmasker/masker";

describe("secret masker", () => {
  const sandbox = sinon.createSandbox();
  afterEach(async () => {
    sandbox.restore();
  });
  describe("dictMatcher", () => {
    it("exact", async () => {
      const output = dictMatcher.match("'world'");
      assert.equal(output, "exact");
    });
    it("none", async () => {
      const output = dictMatcher.match("wersdfw");
      assert.equal(output, "none");
    });
  });

  describe("secretMasker", () => {
    it("not contain", async () => {
      const output = secretMasker.maskSecret("Successfully ran target precommit for project.");
      assert.equal(output, "Successfully ran target precommit for project.");
    });
    it("white list", async () => {
      const output = secretMasker.maskSecret("mysql -p password123456", undefined, [
        "password123456",
      ]);
      assert.equal(output, "mysql -p password123456");
    });
  });
});

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs from "fs-extra";
import sinon from "sinon";
import { expect } from "./utils";

const sandbox = sinon.createSandbox();

beforeEach(() => {
  sandbox.stub(fs, "readJsonSync").returns({ version: "2.0.0" });
});

afterEach(() => {
  sandbox.restore();
});

it("getVersion", async () => {
  const utils = require("../../src/utils");
  utils.version = undefined;
  const version = utils.getVersion();
  expect(version).equals("2.0.0");
});

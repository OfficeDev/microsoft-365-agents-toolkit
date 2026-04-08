// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import "mocha";
import fs from "fs-extra";
import sinon from "sinon";
import properLock from "proper-lockfile";
import { globalStateGet, globalStateUpdate } from "../../src/common/globalState";

describe("Global State Get/Update", () => {
  const sandbox = sinon.createSandbox();

  afterEach(async () => {
    sandbox.restore();
  });

  it("returns stored value if key has been updated before", async () => {
    sandbox.stub<any, any>(fs, "readJSON").resolves({ test: false });
    sandbox.stub<any, any>(fs, "pathExistsSync").callsFake(() => {
      return true;
    });
    sandbox.stub<any, any>(fs, "existsSync").callsFake(() => {
      return true;
    });
    sandbox.stub(properLock, "lock").resolves();
    sandbox.stub(properLock, "unlock").resolves();
    const data = await globalStateGet("test", true);
    assert.strictEqual(data, false);
  });

  it("returns default value if key hasn't been updated before", async () => {
    sandbox.stub<any, any>(fs, "readJSON").resolves({});
    sandbox.stub<any, any>(fs, "pathExistsSync").callsFake(() => {
      return true;
    });
    sandbox.stub<any, any>(fs, "existsSync").callsFake(() => {
      return true;
    });
    sandbox.stub(properLock, "lock");
    sandbox.stub(properLock, "unlock");
    const data = await globalStateGet("test", true);
    assert.strictEqual(data, true);
  });

  it("stores value if globalStateUpdate is called", async () => {
    sandbox.stub<any, any>(fs, "readJSONSync").callsFake(() => {
      return {};
    });
    sandbox.stub<any, any>(fs, "readJSON").callsFake(async () => {
      return {};
    });
    sandbox.stub<any, any>(fs, "pathExistsSync").callsFake(() => {
      return false;
    });
    sandbox.stub<any, any>(fs, "mkdirpSync");
    sandbox.stub<any, any>(fs, "existsSync").callsFake(() => {
      return true;
    });
    sandbox.stub<any, any>(fs, "writeJson").callsFake(async (_file: unknown, object: unknown) => {
      data = object;
    });
    sandbox.stub<any, any>(fs, "writeJsonSync").callsFake((_file: unknown, object: unknown) => {
      data = object;
    });
    sandbox.stub(properLock, "lock");
    sandbox.stub(properLock, "unlock");
    let data: any;
    await globalStateUpdate("test", true);
    assert.deepEqual(data, { test: true });
  });
});

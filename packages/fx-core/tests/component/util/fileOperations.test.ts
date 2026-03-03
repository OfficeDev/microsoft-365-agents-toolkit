// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";
import * as sinon from "sinon";
import fs from "fs-extra";
import * as chai from "chai";
import { zipFolderAsync } from "../../../src/component/utils/fileOperation";
import ignore from "ignore";
import { CacheFileInUse, DeployEmptyFolderError, ZipFileError } from "../../../src";
import * as os from "os";
import * as uuid from "uuid";
import * as path from "path";
import AdmZip from "adm-zip";
import proxyquire from "proxyquire";

describe("Test", () => {
  const sandbox = sinon.createSandbox();
  const tmp = `${os.tmpdir()}/${uuid.v4()}`;
  const tmpFile = `${tmp}/test.txt`;

  class EError extends Error {
    code: string;
    constructor(error: Error) {
      super(error.message);
      this.code = error.message;
    }
  }

  before(async () => {
    await fs.mkdirs(tmp);
    await fs.writeFile(tmpFile, "test");
  });

  after(async () => {
    await fs.remove(tmpFile);
    await fs.rmdir(tmp);
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("should throw error when EBUSY", async () => {
    const err = new EError(new Error("EBUSY"));
    sandbox.stub(fs, "remove").throws(err);
    await zipFolderAsync(tmp, tmpFile, ignore()).catch((e) => {
      chai.expect(e instanceof CacheFileInUse).to.equal(true);
    });
  });

  it("should throw error when Other error", async () => {
    sandbox.stub(fs, "remove").throws(new Error("Other"));
    await zipFolderAsync(tmp, tmpFile, ignore()).catch((e) => {
      chai.expect(e.message).to.equal("Other");
    });
  });

  it("should throw error when folder is empty", async () => {
    const empty = `${os.tmpdir()}/empty`;
    await fs.mkdirs(empty);
    await zipFolderAsync(empty, `./${uuid.v4()}`, ignore()).catch((e) => {
      chai.expect(e instanceof DeployEmptyFolderError).to.equal(true);
    });
    await fs.rmdir(empty);
  });

  it("write to zip throws ERR_OUT_OF_RANGE", async () => {
    const test = new AdmZip();
    const err = new EError(new Error("ERR_OUT_OF_RANGE"));
    sandbox.stub(test, "writeZip").yields(err);
    const zipFolderAsync = proxyquire("../../../src/component/utils/fileOperation", {
      "adm-zip": function() { return test; } // This function will be called instead of `new AdmZip()`
    }).zipFolderAsync;
    await zipFolderAsync(tmp, path.join(tmp, "tmp.zip"), ignore()).catch((e: Error) => {
      chai.expect(e instanceof ZipFileError).to.equal(true);
    });
  });
});

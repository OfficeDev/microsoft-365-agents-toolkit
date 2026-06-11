/* eslint-disable prettier/prettier */
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as chai from "chai";
import fs from "fs-extra";
import ignore from "ignore";
import * as os from "os";
import * as path from "path";
import * as sinon from "sinon";
import * as uuid from "uuid";
import { CacheFileInUse, DeployEmptyFolderError, ZipFileError } from "../../../src";
import { fileOperationDeps, zipFolderAsync } from "../../../src/component/utils/fileOperation";

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
    sandbox.stub(fileOperationDeps, "existsSync").returns(true);
    sandbox.stub(fileOperationDeps, "remove").rejects(err);
    await zipFolderAsync(tmp, tmpFile, ignore()).catch((e) => {
      chai.expect(e instanceof CacheFileInUse).to.equal(true);
    });
  });

  it("should throw error when Other error", async () => {
    sandbox.stub(fileOperationDeps, "existsSync").returns(true);
    sandbox.stub(fileOperationDeps, "remove").rejects(new Error("Other"));
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
    const err = new EError(new Error("ERR_OUT_OF_RANGE"));
    sandbox.stub(fileOperationDeps, "writeZip").rejects(err);
    await zipFolderAsync(tmp, path.join(tmp, "tmp.zip"), ignore()).catch((e: Error) => {
      chai.expect(e instanceof ZipFileError).to.equal(true);
    });
  });

  it("fileOperationDeps.writeZip should reject on callback error", async () => {
    const fakeZip = {
      writeZip: (_cache: string, cb: (err?: Error) => void) => cb(new Error("zip-failed")),
    } as any;

    await fileOperationDeps.writeZip(fakeZip, path.join(tmp, "tmp.zip")).catch((e: Error) => {
      chai.expect(e.message).to.equal("zip-failed");
    });
  });
});

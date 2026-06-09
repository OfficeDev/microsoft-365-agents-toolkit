// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author huajiezhang@microsoft.com
 */

import * as chai from "chai";
import fse from "fs-extra";
import * as sinon from "sinon";
import {
  HelperMethods,
  helperMethodsDeps,
} from "../../src/component/generator/officeAddin/helperMethods";

describe("Generator related Utils", function () {
  describe("fetchAndUnzip", async () => {
    const sandbox = sinon.createSandbox();
    class ZipEntry {
      isDirectory: boolean;
      entryName: string;
      getData() {
        return undefined;
      }
      constructor(isDir: boolean, entryName: string) {
        this.isDirectory = isDir;
        this.entryName = entryName;
      }
    }

    class MockAdmZip {
      getEntries() {
        return [
          new ZipEntry(true, "dir/"),
          new ZipEntry(true, "dir/subdir/"),
          new ZipEntry(false, "dir/subdir/file"),
        ];
      }
    }

    afterEach(() => {
      sandbox.restore();
    });

    it("happy path", async () => {
      sandbox.stub(helperMethodsDeps, "fetchZipFromUrl").resolves(new MockAdmZip() as any);
      const stub1 = sandbox.stub(fse, "ensureDir").resolves();
      const stub2 = sandbox.stub(fse, "writeFile").resolves();
      const res = await HelperMethods.fetchAndUnzip("test", "url", "dest");
      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(stub1.calledOnce);
      chai.assert.isTrue(stub2.calledOnce);
    });

    it("fail case: fetch zip throw error", async () => {
      sandbox.stub(helperMethodsDeps, "fetchZipFromUrl").rejects(new Error());
      const res = await HelperMethods.fetchAndUnzip("test", "url", "dest");
      chai.assert.isTrue(res.isErr());
    });

    it("fail case: fetch zip returns undefined", async () => {
      sandbox.stub(helperMethodsDeps, "fetchZipFromUrl").resolves(undefined);
      const res = await HelperMethods.fetchAndUnzip("test", "url", "dest");
      chai.assert.isTrue(res.isErr());
    });

    it("fail case: ensureDir throws error", async () => {
      sandbox.stub(helperMethodsDeps, "fetchZipFromUrl").resolves(new MockAdmZip() as any);
      sandbox.stub(fse, "ensureDir").rejects(new Error());
      const res = await HelperMethods.fetchAndUnzip("test", "url", "dest");
      chai.assert.isTrue(res.isErr());
    });
  });
});

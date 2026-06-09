// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import * as fs from "fs-extra";
import * as os from "os";
import * as path from "path";
import {
  bundledFloorDir,
  bundledFloorFrom,
  loadBundledFloor,
} from "../../../src/v4/distribution/bundledFloor";
import { computeDigest } from "../../../src/v4/distribution/templateSource";

describe("bundledFloor (v4)", () => {
  describe("bundledFloorFrom (pure)", () => {
    it("computes the digest from the bytes (not baked, decision #6)", () => {
      const bytes = Buffer.from("floor-bytes");
      const floor = bundledFloorFrom("6.11.0", bytes, "/some/templates.zip");
      assert.deepEqual(floor, {
        version: "6.11.0",
        digest: computeDigest(bytes),
        location: "/some/templates.zip",
      });
    });
  });

  describe("bundledFloorDir (location contract)", () => {
    it("resolves the baked floor under the templates folder's v4 subdirectory", () => {
      assert.strictEqual(path.basename(bundledFloorDir()), "v4");
    });
  });

  describe("loadBundledFloor (IO)", () => {
    let dir: string;

    beforeEach(() => {
      dir = fs.mkdtempSync(path.join(os.tmpdir(), "v4-floor-"));
    });

    afterEach(() => {
      fs.removeSync(dir);
    });

    it("reads floor.json + templates.zip and computes the digest", () => {
      const bytes = Buffer.from("the-floor-package");
      fs.writeJsonSync(path.join(dir, "floor.json"), { version: "6.11.0" });
      fs.writeFileSync(path.join(dir, "templates.zip"), bytes);

      const floor = loadBundledFloor(dir);
      assert.strictEqual(floor.version, "6.11.0");
      assert.strictEqual(floor.digest, computeDigest(bytes));
      assert.strictEqual(floor.location, path.join(dir, "templates.zip"));
    });

    it("throws BundledFloorMissing when the manifest is absent", () => {
      assert.throws(() => loadBundledFloor(dir), /BundledFloorMissing|manifest is missing/);
    });

    it("throws BundledFloorMissing when the package zip is absent", () => {
      fs.writeJsonSync(path.join(dir, "floor.json"), { version: "6.11.0" });
      assert.throws(() => loadBundledFloor(dir), /BundledFloorMissing|package is missing/);
    });

    it("throws BundledFloorMalformed when version is missing", () => {
      fs.writeJsonSync(path.join(dir, "floor.json"), { notVersion: "x" });
      fs.writeFileSync(path.join(dir, "templates.zip"), Buffer.from("z"));
      assert.throws(() => loadBundledFloor(dir), /BundledFloorMalformed|no string "version"/);
    });
  });
});

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import {
  computeBundled,
  computeRange,
  computeV4TemplateConfig,
} from "../../../src/v4/distribution/templateConfig";

describe("templateConfig (v4 build-time)", () => {
  describe("computeBundled (bundled = !goproduct)", () => {
    it("internal test build (goproduct=false) → bundled", () => {
      assert.strictEqual(computeBundled(false), true);
    });
    it("shipping build (goproduct=true) → online", () => {
      assert.strictEqual(computeBundled(true), false);
    });
  });

  describe("computeRange", () => {
    it("keeps the previous range when the version still intersects it", () => {
      assert.strictEqual(computeRange("6.10.5", "~6.10"), "~6.10");
    });

    it("widens to ~major.minor when the version no longer intersects", () => {
      assert.strictEqual(computeRange("6.11.0", "~6.10"), "~6.11");
    });

    it("derives ~major.minor from a prerelease's stable target", () => {
      // templates has no odd/even-minor split, so the rc shares the stable range.
      assert.strictEqual(computeRange("6.11.0-rc.0", "~6.10"), "~6.11");
    });

    it("bumps the major when crossing a major boundary", () => {
      assert.strictEqual(computeRange("7.0.0", "~6.10"), "~7.0");
    });

    it("throws on a non-SemVer version (no silent fallback)", () => {
      assert.throws(() => computeRange("not-semver", "~6.10"), /not valid SemVer/);
    });
  });

  describe("computeV4TemplateConfig", () => {
    it("internal rc test build → bundled floor, range from stable target, exact localVersion", () => {
      const config = computeV4TemplateConfig({
        version: "6.11.0-rc.0",
        goproduct: false,
        previousRange: "~6.10",
      });
      assert.deepEqual(config, {
        range: "~6.11",
        bundled: true,
        localVersion: "6.11.0-rc.0",
      });
    });

    it("stable shipping build → online channel, widened range", () => {
      const config = computeV4TemplateConfig({
        version: "6.11.0",
        goproduct: true,
        previousRange: "~6.10",
      });
      assert.deepEqual(config, {
        range: "~6.11",
        bundled: false,
        localVersion: "6.11.0",
      });
    });

    it("prerelease shipped to marketplace pre-release channel → online (bundled=false)", () => {
      const config = computeV4TemplateConfig({
        version: "6.11.0-rc.0",
        goproduct: true,
        previousRange: "~6.11",
      });
      assert.deepEqual(config, {
        range: "~6.11",
        bundled: false,
        localVersion: "6.11.0-rc.0",
      });
    });

    it("patch within the current range keeps the range stable (reproducibility)", () => {
      const config = computeV4TemplateConfig({
        version: "6.10.2",
        goproduct: true,
        previousRange: "~6.10",
      });
      assert.deepEqual(config, {
        range: "~6.10",
        bundled: false,
        localVersion: "6.10.2",
      });
    });
  });
});

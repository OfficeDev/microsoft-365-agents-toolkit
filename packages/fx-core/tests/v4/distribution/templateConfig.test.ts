// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import {
  computeBundled,
  computeRange,
  computeV4PublishVersion,
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

  describe("computeV4PublishVersion", () => {
    it("odd-minor preview → stamps the build date into the patch, drops the suffix", () => {
      assert.strictEqual(computeV4PublishVersion("6.11.1-beta.2026061609.0"), "6.11.2026061609");
    });

    it("odd-minor without a date stamp → major.minor.patch", () => {
      assert.strictEqual(computeV4PublishVersion("6.11.5"), "6.11.5");
    });

    it("even-minor stable → major.minor.patch as-is", () => {
      assert.strictEqual(computeV4PublishVersion("6.10.3"), "6.10.3");
    });

    it("even-minor suffixed (preview minted on stable branch) → stripped (publish-time guard refuses it separately)", () => {
      assert.strictEqual(computeV4PublishVersion("6.10.3-beta.2026061609.0"), "6.10.3");
    });

    it("throws on a non-SemVer version (no silent fallback)", () => {
      assert.throws(() => computeV4PublishVersion("not-semver"), /not valid SemVer/);
    });
  });

  describe("computeRange", () => {
    it("keeps the previous range when the version still intersects it", () => {
      assert.strictEqual(computeRange("6.10.5", "~6.10"), "~6.10");
    });

    it("widens to ~major.minor when the version no longer intersects", () => {
      assert.strictEqual(computeRange("6.11.0", "~6.10"), "~6.11");
    });

    it("derives ~major.minor from a prerelease's minor", () => {
      // Odd-minor prerelease lives on its own minor; range follows that minor.
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
    it("odd-minor preview shipping (goproduct=true) → online, clean date-stamped localVersion", () => {
      const config = computeV4TemplateConfig({
        version: "6.11.1-beta.2026061609.0",
        goproduct: true,
        previousRange: "~6.10",
      });
      assert.deepEqual(config, {
        range: "~6.11",
        bundled: false,
        localVersion: "6.11.2026061609",
      });
    });

    it("even-minor stable shipping (goproduct=true) → online, version as-is", () => {
      const config = computeV4TemplateConfig({
        version: "6.10.3",
        goproduct: true,
        previousRange: "~6.10",
      });
      assert.deepEqual(config, {
        range: "~6.10",
        bundled: false,
        localVersion: "6.10.3",
      });
    });

    it("internal test build (goproduct=false) → bundled floor regardless of suffix", () => {
      const config = computeV4TemplateConfig({
        version: "6.11.1-beta.2026061609.0",
        goproduct: false,
        previousRange: "~6.10",
      });
      assert.deepEqual(config, {
        range: "~6.11",
        bundled: true,
        localVersion: "6.11.2026061609",
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

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { expect } from "chai";
import { generatePlaceholderPng } from "../../../../src/component/generator/openPlugin/placeholderPng";

const PNG_SIGNATURE = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

describe("openPlugin.generatePlaceholderPng", () => {
  it("generates a valid PNG for a 1x1 image", () => {
    const buf = generatePlaceholderPng(1, 255, 0, 0);
    expect(buf.subarray(0, 8).equals(PNG_SIGNATURE)).to.equal(true);
    expect(buf.length).to.be.greaterThan(8);
  });

  it("throws for size 0", () => {
    expect(() => generatePlaceholderPng(0, 0, 0, 0)).to.throw(/positive integer/);
  });

  it("throws for negative size", () => {
    expect(() => generatePlaceholderPng(-5, 0, 0, 0)).to.throw(/positive integer/);
  });

  it("throws for non-integer size", () => {
    expect(() => generatePlaceholderPng(1.5, 0, 0, 0)).to.throw(/positive integer/);
  });
});

/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it } from "mocha";
import { createCliContext } from "../../src/context";

describe("createCliContext()", () => {
  it("should create an AtkContext with default projectPath (cwd)", () => {
    const ctx = createCliContext();
    expect(ctx.projectPath).to.equal(process.cwd());
  });

  it("should accept a custom projectPath", () => {
    const ctx = createCliContext("/tmp/my-project");
    expect(ctx.projectPath).to.equal("/tmp/my-project");
  });

  it("should generate a unique correlationId", () => {
    const ctx1 = createCliContext();
    const ctx2 = createCliContext();
    expect(ctx1.correlationId).to.be.a("string");
    expect(ctx1.correlationId).to.not.equal(ctx2.correlationId);
  });

  it("should have logger, telemetry, ui, and auth wired", () => {
    const ctx = createCliContext();
    expect(ctx.logger).to.exist;
    expect(ctx.telemetry).to.exist;
    expect(ctx.ui).to.exist;
    expect(ctx.auth).to.exist;
  });
});

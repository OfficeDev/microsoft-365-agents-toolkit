/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it } from "mocha";
import { createAtkContext } from "../../../src/core/context";
import { createMockContext } from "../testHelper";

describe("AtkContext", () => {
  describe("createAtkContext()", () => {
    it("should create context with all required fields", () => {
      const mock = createMockContext();
      const ctx = createAtkContext({
        auth: mock.auth,
        logger: mock.logger,
        telemetry: mock.telemetry,
        ui: mock.ui,
      });
      expect(ctx.auth).to.equal(mock.auth);
      expect(ctx.logger).to.equal(mock.logger);
      expect(ctx.telemetry).to.equal(mock.telemetry);
      expect(ctx.ui).to.equal(mock.ui);
      expect(ctx.correlationId).to.be.a("string").and.not.be.empty;
    });

    it("should use provided correlation ID", () => {
      const mock = createMockContext();
      const ctx = createAtkContext({
        auth: mock.auth,
        logger: mock.logger,
        telemetry: mock.telemetry,
        ui: mock.ui,
        correlationId: "custom-id",
      });
      expect(ctx.correlationId).to.equal("custom-id");
    });

    it("should generate UUID-formatted correlation ID when not provided", () => {
      const mock = createMockContext();
      const ctx = createAtkContext({
        auth: mock.auth,
        logger: mock.logger,
        telemetry: mock.telemetry,
        ui: mock.ui,
      });
      expect(ctx.correlationId).to.match(
        /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/
      );
    });
  });
});

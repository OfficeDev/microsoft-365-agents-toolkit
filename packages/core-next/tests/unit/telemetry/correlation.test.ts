/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it } from "mocha";
import { correlationScope, getCurrentCorrelationId } from "../../../src/telemetry/correlation";

describe("Correlation", () => {
  describe("correlationScope()", () => {
    it("should make correlation ID available inside the scope", () => {
      const result = correlationScope("scope-id-1", () => {
        return getCurrentCorrelationId();
      });
      expect(result).to.equal("scope-id-1");
    });

    it("should return empty string outside any scope", () => {
      expect(getCurrentCorrelationId()).to.equal("");
    });

    it("should support nested scopes with inner overriding outer", () => {
      correlationScope("outer", () => {
        expect(getCurrentCorrelationId()).to.equal("outer");
        correlationScope("inner", () => {
          expect(getCurrentCorrelationId()).to.equal("inner");
        });
        expect(getCurrentCorrelationId()).to.equal("outer");
      });
    });
  });
});

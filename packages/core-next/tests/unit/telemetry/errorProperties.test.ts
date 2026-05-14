/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it } from "mocha";
import { extractErrorProperties } from "../../../src/telemetry/errorProperties";
import { userError, systemError } from "../../../src/core/error";
import { TelemetryProperty } from "../../../src/telemetry/types";

describe("extractErrorProperties()", () => {
  it("should extract code, kind, and message", () => {
    const error = userError("BadInput", "invalid value");
    const props = extractErrorProperties(error);
    expect(props[TelemetryProperty.ErrorCode]).to.equal("BadInput");
    expect(props[TelemetryProperty.ErrorKind]).to.equal("user");
    expect(props[TelemetryProperty.ErrorMessage]).to.equal("invalid value");
  });

  it("should include source if present", () => {
    const error = systemError("Timeout", "timed out", { source: "http" });
    const props = extractErrorProperties(error);
    expect(props[TelemetryProperty.Component]).to.equal("http");
  });

  it("should include inner error message if present", () => {
    const inner = new Error("root cause");
    const error = systemError("Wrapper", "wrapped", { inner });
    const props = extractErrorProperties(error);
    expect(props["inner-error-message"]).to.equal("root cause");
  });

  it("should not include component or inner when absent", () => {
    const error = userError("Simple", "simple error");
    const props = extractErrorProperties(error);
    expect(props).to.not.have.property(TelemetryProperty.Component);
    expect(props).to.not.have.property("inner-error-message");
  });
});

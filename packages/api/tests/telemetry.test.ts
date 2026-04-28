// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import "mocha";
import chai from "chai";
import spies from "chai-spies";
import { TelemetryReporter } from "../src/utils/telemetry";

chai.use(spies);
const expect = chai.expect;

const pluginName = "The way that can be told of is not an unvarying way";
const sandbox = chai.spy.sandbox();

class MockTelemetryReporter implements TelemetryReporter {
  sendTelemetryErrorEvent(
    {}: string,
    {}: { [p: string]: string },
    {}: { [p: string]: number },
    {}: string[]
  ): void {
    // do nothing
  }

  sendTelemetryEvent({}: string, {}: { [p: string]: string }, {}: { [p: string]: number }): void {
    // do nothing
  }

  sendTelemetryException(
    {}: Error,
    {}: { [p: string]: string },
    {}: { [p: string]: number }
  ): void {
    // do nothing
  }
}

describe("telemetry", () => {
  let reporter: TelemetryReporter;

  beforeEach(() => {
    reporter = new MockTelemetryReporter();
    sandbox.on(reporter, [
      "sendTelemetryErrorEvent",
      "sendTelemetryEvent",
      "sendTelemetryException",
    ]);
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("invoke times", () => {
    reporter.sendTelemetryEvent(
      "sampleEvent",
      { stringProp: "some string" },
      { numericMeasure: 123 }
    );

    expect(reporter.sendTelemetryEvent).to.have.been.called.once;
  });
});

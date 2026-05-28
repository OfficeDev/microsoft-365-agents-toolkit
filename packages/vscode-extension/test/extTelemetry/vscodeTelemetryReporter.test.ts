/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-empty-function */
/* eslint-disable @typescript-eslint/no-var-requires */
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// eslint-disable-next-line import/default
import TelemetryReporter from "@vscode/extension-telemetry";
import * as sinon from "sinon";
import { VSCodeTelemetryReporter } from "../../src/telemetry/vscodeTelemetryReporter";
import { MockTelemetryReporter } from "../mocks/mockTools";
import { featureFlagManager } from "@microsoft/teamsfx-core";

const featureFlags = featureFlagManager.listEnabled().join(";") ?? "";

describe("vscodeTelemetryReporter", () => {
  let tester: VSCodeTelemetryReporter;
  const sandbox = sinon.createSandbox();
  const reporterStub = new MockTelemetryReporter();
  let sendTelemetryEventSpy: sinon.SinonSpy;
  let sendTelemetryExceptionSpy: sinon.SinonSpy;
  let sendTelemetryErrorEventSpy: sinon.SinonSpy;

  beforeEach(() => {
    tester = new VSCodeTelemetryReporter(
      "test",
      "1.0.0-rc.1",
      "test",
      reporterStub as unknown as TelemetryReporter
    );
    tester.addSharedProperty("project-id", "");
    tester.addSharedProperty("programming-language", "");
    tester.addSharedProperty("host-type", "");
    tester.addSharedProperty("is-from-sample", "");

    sendTelemetryEventSpy = sandbox.spy(reporterStub, "sendTelemetryEvent");
    sendTelemetryExceptionSpy = sandbox.spy(reporterStub, "sendTelemetryException");
    sendTelemetryErrorEventSpy = sandbox.spy(reporterStub, "sendTelemetryErrorEvent");
  });

  afterEach(() => {
    tester.dispose();
    sandbox.restore();
  });

  it("sendTelemetryEvent", () => {
    tester.sendTelemetryEvent(
      "sampleEvent",
      { stringProp: "some string" },
      { numericMeasure: 123 }
    );

    sinon.assert.calledOnceWithMatch(
      sendTelemetryEventSpy,
      "sampleEvent",
      {
        stringProp: "some string",
        "project-id": "",
        "correlation-id": "",
        "feature-flags": featureFlags,
        "programming-language": "",
        "host-type": "",
        "is-from-sample": "",
      },
      { numericMeasure: 123 }
    );
  });

  it("sendTelemetryErrorEvent", () => {
    tester.sendTelemetryErrorEvent(
      "sampleErrorEvent",
      {
        stringProp: "some string",
        "error-stack": "some user stack trace at (C:/fake_path/fake_file:1:1)",
      },
      { numericMeasure: 123 },
      ["error-stack"]
    );

    sinon.assert.calledOnceWithMatch(
      sendTelemetryErrorEventSpy,
      "sampleErrorEvent",
      {
        stringProp: "some string",
        "error-stack": "some user stack trace at (<REDACTED: user-file-path>/fake_file:1:1)",
        "project-id": "",
        "correlation-id": "",
        "feature-flags": featureFlags,
        "programming-language": "",
        "host-type": "",
        "is-from-sample": "",
      },
      { numericMeasure: 123 }
    );
  });

  it("sendTelemetryErrorEvent: not overwrite correlationId if existing", () => {
    tester.sendTelemetryErrorEvent(
      "sampleErrorEvent",
      {
        stringProp: "some string",
        "error-stack": "some user stack trace at (C:/fake_path/fake_file:1:1)",
        "correlation-id": "fakeId",
      },
      { numericMeasure: 123 },
      ["error-stack"]
    );

    sinon.assert.calledOnceWithMatch(
      sendTelemetryErrorEventSpy,
      "sampleErrorEvent",
      {
        stringProp: "some string",
        "error-stack": "some user stack trace at (<REDACTED: user-file-path>/fake_file:1:1)",
        "project-id": "",
        "correlation-id": "fakeId",
        "feature-flags": featureFlags,
        "programming-language": "",
        "host-type": "",
        "is-from-sample": "",
      },
      { numericMeasure: 123 }
    );
  });

  it("sendTelemetryException", () => {
    const error = new Error("error for test");
    tester.sendTelemetryException(error, { stringProp: "some string" }, { numericMeasure: 123 });

    sinon.assert.calledOnceWithMatch(
      sendTelemetryExceptionSpy,
      error,
      {
        stringProp: "some string",
        "project-id": "",
        "correlation-id": "",
        "feature-flags": featureFlags,
        "programming-language": "",
        "host-type": "",
        "is-from-sample": "",
      },
      { numericMeasure: 123 }
    );
  });
});

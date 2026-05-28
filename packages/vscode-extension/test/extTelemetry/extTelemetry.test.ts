import { Stage, UserError } from "@microsoft/teamsfx-api";
import { maskSecret, telemetryUtils } from "@microsoft/teamsfx-core";
import * as globalState from "@microsoft/teamsfx-core/build/common/globalState";
import * as chai from "chai";
import fs from "fs-extra";
import * as sinon from "sinon";
import { Uri } from "vscode";
import * as globalVariables from "../../src/globalVariables";
import * as telemetryModule from "../../src/telemetry/extTelemetry";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { TelemetryEvent } from "../../src/telemetry/extTelemetryEvents";
import * as vscTelemetryUtils from "../../src/utils/telemetryUtils";
import { MockTelemetryReporter } from "../mocks/mockTools";

describe("ExtTelemetry", () => {
  chai.util.addProperty(ExtTelemetry, "reporter", () => {});
  let sendTelemetryErrorEventSpy: sinon.SinonSpy<
    [
      eventName: string,
      properties?: { [key: string]: string } | undefined,
      measurements?: { [key: string]: number } | undefined,
      errorProps?: string[] | undefined
    ],
    void
  >;
  let sendTelemetryEventSpy: sinon.SinonSpy<
    [
      eventName: string,
      properties?: { [key: string]: string } | undefined,
      measurements?: { [key: string]: number } | undefined
    ],
    void
  >;
  let sendTelemetryExceptionSpy: sinon.SinonSpy<
    [
      error: Error,
      properties?: { [key: string]: string } | undefined,
      measurements?: { [key: string]: number } | undefined
    ],
    void
  >;

  describe("setHasSentTelemetry", () => {
    it("query-expfeature", () => {
      const eventName = "query-expfeature";
      ExtTelemetry.hasSentTelemetry = false;
      ExtTelemetry.setHasSentTelemetry(eventName);
      chai.expect(ExtTelemetry.hasSentTelemetry).equals(false);
    });

    it("other-event", () => {
      const eventName = "other-event";
      ExtTelemetry.hasSentTelemetry = false;
      ExtTelemetry.setHasSentTelemetry(eventName);
      chai.expect(ExtTelemetry.hasSentTelemetry).equals(true);
    });
  });

  describe("stageToEvent", () => {
    it("Stage.create", () => {
      const stage = Stage.create;
      chai.expect(ExtTelemetry.stageToEvent(stage)).equals(TelemetryEvent.CreateProject);
    });

    it("Stage.provision", () => {
      const stage = Stage.provision;
      chai.expect(ExtTelemetry.stageToEvent(stage)).equals(TelemetryEvent.Provision);
    });

    it("Stage.deploy", () => {
      const stage = Stage.deploy;
      chai.expect(ExtTelemetry.stageToEvent(stage)).equals(TelemetryEvent.Deploy);
    });

    it("Stage.publish", () => {
      const stage = Stage.publish;
      chai.expect(ExtTelemetry.stageToEvent(stage)).equals(TelemetryEvent.Publish);
    });

    it("Stage.creatEnv", () => {
      const stage = Stage.createEnv;
      chai.expect(ExtTelemetry.stageToEvent(stage)).equals(TelemetryEvent.CreateNewEnvironment);
    });

    it("Stage.addWebpart", () => {
      const stage = Stage.addWebpart;
      chai.expect(ExtTelemetry.stageToEvent(stage)).equals(TelemetryEvent.AddWebpart);
    });

    it("Stage.copilotPluginAddAPI", () => {
      const stage = Stage.copilotPluginAddAPI;
      chai.expect(ExtTelemetry.stageToEvent(stage)).equals(TelemetryEvent.CopilotPluginAddAPI);
    });

    it("Stage.syncManifest", () => {
      const stage = Stage.syncManifest;
      chai.expect(ExtTelemetry.stageToEvent(stage)).equals(TelemetryEvent.SyncManifest);
    });

    it("unknown", () => {
      const stage = "unknown";
      chai.expect(ExtTelemetry.stageToEvent(stage as Stage)).equals(undefined);
    });
  });

  describe("Send Telemetry", () => {
    const sandbox = sinon.createSandbox();
    const reporterStub = new MockTelemetryReporter();

    beforeEach(() => {
      sendTelemetryErrorEventSpy = sandbox.spy(reporterStub, "sendTelemetryErrorEvent");
      sendTelemetryEventSpy = sandbox.spy(reporterStub, "sendTelemetryEvent");
      sendTelemetryExceptionSpy = sandbox.spy(reporterStub, "sendTelemetryException");
      sandbox.stub(ExtTelemetry, "reporter").value(reporterStub);
      sandbox.stub(ExtTelemetry, "settingsVersion").value("1.0.0");
      sandbox.stub(fs, "pathExistsSync").returns(false);
      sandbox.stub(globalVariables, "workspaceUri").value(Uri.file("test"));
      sandbox.stub(globalVariables, "isSPFxProject").value(false);
      sandbox.stub(globalVariables, "isExistingUser").value("no");
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("sendTelemetryEvent", () => {
      ExtTelemetry.sendTelemetryEvent(
        "sampleEvent",
        { stringProp: "some string" },
        { numericMeasure: 123 }
      );

      sinon.assert.calledOnceWithMatch(
        sendTelemetryEventSpy,
        "sampleEvent",
        {
          stringProp: "some string",
          component: "extension",
          "is-existing-user": "no",
          "is-spfx": "false",
          "settings-version": "1.0.0",
        },
        { numericMeasure: 123 }
      );
    });

    it("sendTelemetryErrorEvent", () => {
      const error = new UserError(
        "test",
        "UserTestError",
        "test error message",
        "displayed test error message"
      );
      ExtTelemetry.sendTelemetryErrorEvent(
        "sampleEvent",
        error,
        { stringProp: "some string" },
        { numericMeasure: 123 },
        ["errorProps"]
      );

      sinon.assert.calledOnceWithMatch(
        sendTelemetryErrorEventSpy,
        "sampleEvent",
        {
          stringProp: "some string",
          component: "extension",
          success: "no",
          "is-existing-user": "no",
          "is-spfx": "false",
          "settings-version": "1.0.0",
          "error-type": "user",
          "error-name": "UserTestError",
          "err-message": maskSecret(error.message),
          "err-stack": telemetryUtils.extractMethodNamesFromErrorStack(error.stack),
          "error-code": "test.UserTestError",
          "error-component": "",
          "error-method": "",
          "error-source": "",
          "error-stage": "",
        },
        { numericMeasure: 123 },
        ["errorProps"]
      );
    });

    it("sendTelemetryException", () => {
      const error = new UserError("test", "UserTestError", "test error message");
      ExtTelemetry.sendTelemetryException(
        error,
        { stringProp: "some string" },
        { numericMeasure: 123 }
      );

      sinon.assert.calledOnceWithMatch(
        sendTelemetryExceptionSpy,
        error,
        {
          stringProp: "some string",
          component: "extension",
          "is-existing-user": "no",
          "is-spfx": "false",
          "settings-version": "1.0.0",
        },
        { numericMeasure: 123 }
      );
    });
  });

  describe("deactivate event", () => {
    const sandbox = sinon.createSandbox();
    afterEach(() => {
      sandbox.restore();
    });
    it("cacheTelemetryEventAsync", async () => {
      const clock = sandbox.useFakeTimers();
      let state = "";
      sandbox.stub(telemetryModule, "lastCorrelationId").value("correlation-id");
      sandbox.stub(vscTelemetryUtils, "getProjectId").resolves("project-id");
      const globalStateUpdateStub = sandbox
        .stub(globalState, "globalStateUpdate")
        .callsFake(async (key, value) => (state = value));
      const eventName = "deactivate";

      await ExtTelemetry.cacheTelemetryEventAsync(eventName);

      sandbox.assert.calledOnce(globalStateUpdateStub);
      const telemetryEvents = {
        eventName: eventName,
        properties: {
          "correlation-id": "correlation-id",
          "project-id": "project-id",
          timestamp: new clock.Date().toISOString(),
        },
      };
      const newValue = JSON.stringify(telemetryEvents);
      chai.expect(state).equals(newValue);
      clock.restore();
    });

    it("sendCachedTelemetryEventsAsync", async () => {
      const reporterStub = new MockTelemetryReporter();
      sendTelemetryEventSpy = sandbox.spy(reporterStub, "sendTelemetryEvent");
      sandbox.stub(ExtTelemetry, "reporter").value(reporterStub);
      const timestamp = new Date().toISOString();
      const telemetryEvents = {
        eventName: "deactivate",
        properties: {
          "correlation-id": "correlation-id",
          "project-id": "project-id",
          timestamp: timestamp,
        },
      };
      const telemetryData = JSON.stringify(telemetryEvents);
      sandbox.stub(globalState, "globalStateGet").callsFake(async () => telemetryData);
      sandbox.stub(globalState, "globalStateUpdate");

      await ExtTelemetry.sendCachedTelemetryEventsAsync();

      sinon.assert.calledOnceWithMatch(sendTelemetryEventSpy, "deactivate", {
        "correlation-id": "correlation-id",
        "project-id": "project-id",
        timestamp: timestamp,
      });
    });
  });
});

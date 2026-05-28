import sinon from "ts-sinon";
import * as chai from "chai";
import { OfficeChatTelemetryData } from "../../src/officeChat/telemetry";
import { Correlator } from "@microsoft/teamsfx-core";
import {
  TelemetryProperty,
  TelemetrySuccess,
  TelemetryTriggerFrom,
} from "../../src/telemetry/extTelemetryEvents";
import * as utils from "../../src/chat/utils";
import * as coreTools from "@microsoft/teamsfx-core/build/common/stringUtils";

describe("OfficeChatTelemetryData", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
    OfficeChatTelemetryData.requestData = {};
  });

  it("constructor", () => {
    sandbox.stub(Correlator, "getId").returns("testCorrelationId");
    const officeChatTelemetryData = new OfficeChatTelemetryData(
      "testCommand",
      "testRequestId",
      0,
      "testParticipantId"
    );

    const telemetryDataProperties = officeChatTelemetryData.telemetryData.properties;
    chai.assert.equal(telemetryDataProperties[TelemetryProperty.CopilotChatCommand], "testCommand");
    chai.assert.equal(
      telemetryDataProperties[TelemetryProperty.CopilotChatRequestId],
      "testRequestId"
    );
    chai.assert.equal(
      telemetryDataProperties[TelemetryProperty.TriggerFrom],
      TelemetryTriggerFrom.CopilotChat
    );
    chai.assert.equal(
      telemetryDataProperties[TelemetryProperty.CorrelationId],
      "testCorrelationId"
    );
    chai.assert.equal(
      telemetryDataProperties[TelemetryProperty.CopilotChatParticipantId],
      "testParticipantId"
    );

    chai.assert.equal(officeChatTelemetryData.command, "testCommand");
    chai.assert.equal(officeChatTelemetryData.requestId, "testRequestId");
    chai.assert.equal(officeChatTelemetryData.startTime, 0);
    chai.assert.equal(officeChatTelemetryData.participantId, "testParticipantId");
    chai.assert.equal(officeChatTelemetryData.hasComplete, false);
    chai.assert.equal(officeChatTelemetryData.hostType, "");
    chai.assert.equal(officeChatTelemetryData.relatedSampleName, "");
    chai.assert.equal(officeChatTelemetryData.timeToFirstToken, -1);

    chai.assert.equal(
      OfficeChatTelemetryData.requestData["testRequestId"],
      officeChatTelemetryData
    );
  });

  it("properties", () => {
    sandbox.stub(Correlator, "getId").returns("testCorrelationId");
    const officeChatTelemetryData = new OfficeChatTelemetryData(
      "testCommand",
      "testRequestId",
      0,
      "testParticipantId"
    );

    const properties = officeChatTelemetryData.properties;

    chai.assert.equal(properties[TelemetryProperty.CopilotChatCommand], "testCommand");
    chai.assert.equal(properties[TelemetryProperty.CopilotChatRequestId], "testRequestId");
    chai.assert.equal(properties[TelemetryProperty.TriggerFrom], TelemetryTriggerFrom.CopilotChat);
    chai.assert.equal(properties[TelemetryProperty.CorrelationId], "testCorrelationId");
    chai.assert.equal(properties[TelemetryProperty.CopilotChatParticipantId], "testParticipantId");
    chai.assert.equal(officeChatTelemetryData.hostType, "");
    chai.assert.equal(officeChatTelemetryData.relatedSampleName, "");
    chai.assert.equal(officeChatTelemetryData.timeToFirstToken, -1);
  });

  describe("measurements", () => {
    afterEach(() => {
      sandbox.restore();
      OfficeChatTelemetryData.requestData = {};
    });

    it("after init", () => {
      sandbox.stub(Correlator, "getId").returns("testCorrelationId");
      const officeChatTelemetryData = new OfficeChatTelemetryData(
        "testCommand",
        "testRequestId",
        0,
        "testParticipantId"
      );

      const measurements = officeChatTelemetryData.measurements;

      chai.assert.equal(Object.keys(measurements).length, 0);
    });

    it("after complete", () => {
      sandbox.stub(Correlator, "getId").returns("testCorrelationId");
      sandbox.stub(performance, "now").returns(100);
      sandbox.stub(utils, "countMessagesTokens").returns(200);
      const officeChatTelemetryData = new OfficeChatTelemetryData(
        "testCommand",
        "testRequestId",
        0,
        "testParticipantId"
      );

      officeChatTelemetryData.markComplete();

      const measurements = officeChatTelemetryData.measurements;

      chai.assert.equal(measurements[TelemetryProperty.CopilotChatRequestToken], 200);
      chai.assert.equal(measurements[TelemetryProperty.CopilotChatResponseToken], 200);
      chai.assert.equal(measurements[TelemetryProperty.CopilotChatTimeToComplete], 0.1);
      chai.assert.equal(measurements[TelemetryProperty.CopilotChatTimeToFirstToken], -1);
      chai.assert.equal(measurements[TelemetryProperty.CopilotChatRequestTokenPerSecond], 2000);
      chai.assert.equal(measurements[TelemetryProperty.CopilotChatResponseTokenPerSecond], 2000);
    });
  });

  it("createByParticipant", () => {
    sandbox.stub(performance, "now").returns(100);
    sandbox.stub(coreTools, "getUuid").returns("testRequestId");

    const officeChatTelemetryData = OfficeChatTelemetryData.createByParticipant(
      "testParticipantId",
      "testCommand"
    );

    chai.assert.equal(officeChatTelemetryData.command, "testCommand");
    chai.assert.equal(officeChatTelemetryData.participantId, "testParticipantId");
    chai.assert.equal(officeChatTelemetryData.startTime, 100);
    chai.assert.equal(officeChatTelemetryData.requestId, "testRequestId");
  });

  describe("get", () => {
    afterEach(() => {
      sandbox.restore();
      OfficeChatTelemetryData.requestData = {};
    });

    it("unknow requestId", () => {
      chai.assert.isUndefined(OfficeChatTelemetryData.get("unknowRequestId"));
    });

    it("known requestId", () => {
      sandbox.stub(Correlator, "getId").returns("testCorrelationId");
      const officeChatTelemetryData = new OfficeChatTelemetryData(
        "testCommand",
        "testRequestId",
        0,
        "testParticipantId"
      );

      chai.assert.equal(OfficeChatTelemetryData.get("testRequestId"), officeChatTelemetryData);
    });
  });

  it("extendBy", () => {
    const officeChatTelemetryData = OfficeChatTelemetryData.createByParticipant(
      "testParticipantId",
      "testCommand"
    );

    officeChatTelemetryData.extendBy({ testProperty: "testValue" }, { testMeasurement: 1 });

    chai.assert.equal(officeChatTelemetryData.properties["testProperty"], "testValue");
    chai.assert.equal(officeChatTelemetryData.measurements["testMeasurement"], 1);
  });

  it("markComplete", () => {
    sandbox.stub(utils, "countMessagesTokens").returns(100);
    sandbox.stub(performance, "now").returns(100);
    const officeChatTelemetryData = new OfficeChatTelemetryData(
      "testCommand",
      "testRequestId",
      0,
      "testParticipantId"
    );

    chai.assert.equal(officeChatTelemetryData.hasComplete, false);

    officeChatTelemetryData.markComplete();

    chai.assert.equal(officeChatTelemetryData.hasComplete, true);
    chai.assert.equal(
      officeChatTelemetryData.telemetryData.measurements[TelemetryProperty.CopilotChatRequestToken],
      100
    );
    chai.assert.equal(
      officeChatTelemetryData.telemetryData.measurements[
        TelemetryProperty.CopilotChatResponseToken
      ],
      100
    );
    chai.assert.equal(
      officeChatTelemetryData.telemetryData.measurements[
        TelemetryProperty.CopilotChatTimeToComplete
      ],
      0.1
    );
    chai.assert.equal(
      officeChatTelemetryData.telemetryData.properties[TelemetryProperty.Success],
      TelemetrySuccess.Yes
    );
    chai.assert.equal(
      officeChatTelemetryData.telemetryData.properties[TelemetryProperty.CopilotChatCompleteType],
      "success"
    );
    chai.assert.equal(
      officeChatTelemetryData.telemetryData.properties[TelemetryProperty.HostType],
      ""
    );
    chai.assert.equal(
      officeChatTelemetryData.telemetryData.properties[
        TelemetryProperty.CopilotChatRelatedSampleName
      ],
      ""
    );
    chai.assert.equal(
      officeChatTelemetryData.telemetryData.measurements[
        TelemetryProperty.CopilotChatTimeToFirstToken
      ],
      -1
    );
    chai.assert.equal(
      officeChatTelemetryData.telemetryData.measurements[
        TelemetryProperty.CopilotChatRequestTokenPerSecond
      ],
      1000
    );
    chai.assert.equal(
      officeChatTelemetryData.telemetryData.measurements[
        TelemetryProperty.CopilotChatResponseTokenPerSecond
      ],
      1000
    );

    officeChatTelemetryData.markComplete("fail");
    chai.assert.equal(
      officeChatTelemetryData.telemetryData.properties[TelemetryProperty.CopilotChatCompleteType],
      "success"
    );
  });
});

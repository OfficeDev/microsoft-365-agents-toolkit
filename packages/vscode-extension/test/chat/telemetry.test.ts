import * as chai from "chai";
import { ChatTelemetryData } from "../../src/chat/telemetry";
import {
  TelemetryProperty,
  TelemetrySuccess,
  TelemetryTriggerFrom,
} from "../../src/telemetry/extTelemetryEvents";
import sinon from "ts-sinon";
import { Correlator } from "@microsoft/teamsfx-core";
import * as utils from "../../src/chat/utils";
import * as coreTools from "@microsoft/teamsfx-core/build/common/stringUtils";

describe("ChatTelemetryData", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
    ChatTelemetryData.requestData = {};
  });

  it("constructor", () => {
    sandbox.stub(Correlator, "getId").returns("testCorrelationId");
    const chatTelemetryData = new ChatTelemetryData(
      "testCommand",
      "testRequestId",
      0,
      "testParticipantId"
    );

    const telemetryDataProperties = chatTelemetryData.telemetryData.properties;
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

    chai.assert.equal(chatTelemetryData.command, "testCommand");
    chai.assert.equal(chatTelemetryData.requestId, "testRequestId");
    chai.assert.equal(chatTelemetryData.startTime, 0);
    chai.assert.equal(chatTelemetryData.participantId, "testParticipantId");
    chai.assert.equal(chatTelemetryData.hasComplete, false);

    chai.assert.equal(ChatTelemetryData.requestData["testRequestId"], chatTelemetryData);
  });

  it("properties", () => {
    sandbox.stub(Correlator, "getId").returns("testCorrelationId");
    const chatTelemetryData = new ChatTelemetryData(
      "testCommand",
      "testRequestId",
      0,
      "testParticipantId"
    );

    const properties = chatTelemetryData.properties;

    chai.assert.equal(properties[TelemetryProperty.CopilotChatCommand], "testCommand");
    chai.assert.equal(properties[TelemetryProperty.CopilotChatRequestId], "testRequestId");
    chai.assert.equal(properties[TelemetryProperty.TriggerFrom], TelemetryTriggerFrom.CopilotChat);
    chai.assert.equal(properties[TelemetryProperty.CorrelationId], "testCorrelationId");
    chai.assert.equal(properties[TelemetryProperty.CopilotChatParticipantId], "testParticipantId");
  });

  describe("measurements", () => {
    afterEach(() => {
      sandbox.restore();
      ChatTelemetryData.requestData = {};
    });

    it("after init", () => {
      sandbox.stub(Correlator, "getId").returns("testCorrelationId");
      const chatTelemetryData = new ChatTelemetryData(
        "testCommand",
        "testRequestId",
        0,
        "testParticipantId"
      );

      const measurements = chatTelemetryData.measurements;

      chai.assert.equal(Object.keys(measurements).length, 0);
    });

    it("after complete", () => {
      sandbox.stub(Correlator, "getId").returns("testCorrelationId");
      sandbox.stub(Date, "now").returns(100);
      sandbox.stub(utils, "countMessagesTokens").returns(200);
      const chatTelemetryData = new ChatTelemetryData(
        "testCommand",
        "testRequestId",
        0,
        "testParticipantId"
      );

      chatTelemetryData.markComplete();

      const measurements = chatTelemetryData.measurements;

      chai.assert.equal(measurements[TelemetryProperty.CopilotChatTokenCount], 200);
      chai.assert.equal(measurements[TelemetryProperty.CopilotChatTimeToComplete], 100);
    });
  });

  it("createByParticipant", () => {
    sandbox.stub(Date, "now").returns(100);
    sandbox.stub(coreTools, "getUuid").returns("testRequestId");

    const chatTelemetryData = ChatTelemetryData.createByParticipant(
      "testParticipantId",
      "testCommand"
    );

    chai.assert.equal(chatTelemetryData.command, "testCommand");
    chai.assert.equal(chatTelemetryData.participantId, "testParticipantId");
    chai.assert.equal(chatTelemetryData.startTime, 100);
    chai.assert.equal(chatTelemetryData.requestId, "testRequestId");
  });

  describe("get", () => {
    afterEach(() => {
      sandbox.restore();
      ChatTelemetryData.requestData = {};
    });

    it("unknow requestId", () => {
      chai.assert.isUndefined(ChatTelemetryData.get("unknowRequestId"));
    });

    it("known requestId", () => {
      sandbox.stub(Correlator, "getId").returns("testCorrelationId");
      const chatTelemetryData = new ChatTelemetryData(
        "testCommand",
        "testRequestId",
        0,
        "testParticipantId"
      );

      chai.assert.equal(ChatTelemetryData.get("testRequestId"), chatTelemetryData);
    });
  });

  it("extendBy", () => {
    const chatTelemetryData = ChatTelemetryData.createByParticipant(
      "testParticipantId",
      "testCommand"
    );

    chatTelemetryData.extendBy({ testProperty: "testValue" }, { testMeasurement: 1 });

    chai.assert.equal(chatTelemetryData.properties["testProperty"], "testValue");
    chai.assert.equal(chatTelemetryData.measurements["testMeasurement"], 1);
  });

  it("markComplete", () => {
    sandbox.stub(utils, "countMessagesTokens").returns(100);
    sandbox.stub(Date, "now").returns(100);
    const chatTelemetryData = new ChatTelemetryData(
      "testCommand",
      "testRequestId",
      0,
      "testParticipantId"
    );

    chai.assert.equal(chatTelemetryData.hasComplete, false);

    chatTelemetryData.markComplete();

    chai.assert.equal(chatTelemetryData.hasComplete, true);
    chai.assert.equal(
      chatTelemetryData.telemetryData.measurements[TelemetryProperty.CopilotChatTokenCount],
      100
    );
    chai.assert.equal(
      chatTelemetryData.telemetryData.measurements[TelemetryProperty.CopilotChatTimeToComplete],
      100
    );
    chai.assert.equal(
      chatTelemetryData.telemetryData.properties[TelemetryProperty.Success],
      TelemetrySuccess.Yes
    );
    chai.assert.equal(
      chatTelemetryData.telemetryData.properties[TelemetryProperty.CopilotChatCompleteType],
      "success"
    );

    chatTelemetryData.markComplete("unsupportedPrompt");
    chai.assert.equal(
      chatTelemetryData.telemetryData.properties[TelemetryProperty.CopilotChatCompleteType],
      "success"
    );
  });
});

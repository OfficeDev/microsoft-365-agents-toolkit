import * as chai from "chai";
import * as sinon from "sinon";
import { CancellationToken } from "../mocks/vsc";
import { TeamsChatCommand } from "../../src/chat/consts";
import * as handler from "../../src/chat/handlers";
import { ChatContext, ChatRequest, ChatResponseStream, commands, ChatResultFeedback } from "vscode";
import * as createCommandHandler from "../../src/chat/commands/create/createCommandHandler";
import * as nextStepCommandHandler from "../../src/chat/commands/nextstep/nextstepCommandHandler";
import * as telemetry from "../../src/chat/telemetry";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetryTriggerFrom,
} from "../../src/telemetry/extTelemetryEvents";
import * as util from "../../src/chat/utils";
import { Correlator } from "@microsoft/teamsfx-core";
import { openUrlCommandHandler } from "../../src/chat/handlers";
import { CommandKey } from "../../src/constants";

describe("chat handlers", () => {
  afterEach(() => {
    sinon.restore();
  });

  describe("chatRequestHandler()", () => {
    const sandbox = sinon.createSandbox();
    const response = {
      markdown: sandbox.stub(),
      button: sandbox.stub(),
    };
    const token = new CancellationToken();

    afterEach(async () => {
      sandbox.restore();
    });

    it("call createCommandHandler", async () => {
      const request = {
        prompt: "fakePrompt",
        command: TeamsChatCommand.Create,
        references: [],
        location: 1,
        attempt: 0,
        enableCommandDetection: false,
      } as ChatRequest;
      const createCommandHandlerStub = sandbox.stub(createCommandHandler, "default");
      handler.chatRequestHandler(
        request,
        {} as unknown as ChatContext,
        response as unknown as ChatResponseStream,
        token
      );
      chai
        .expect(
          createCommandHandlerStub.calledOnceWith(
            request,
            {} as unknown as ChatContext,
            response as unknown as ChatResponseStream,
            token
          )
        )
        .to.equal(true);
    });

    it("call nextStepCommandHandler", async () => {
      const request = {
        prompt: "fakePrompt",
        command: TeamsChatCommand.NextStep,
        references: [],
        location: 1,
        attempt: 0,
        enableCommandDetection: false,
      } as ChatRequest;

      const nextStepCommandHandlerStub = sandbox.stub(nextStepCommandHandler, "default");
      handler.chatRequestHandler(
        request,
        {} as unknown as ChatContext,
        response as unknown as ChatResponseStream,
        token
      );
      chai
        .expect(
          nextStepCommandHandlerStub.calledOnceWith(
            request,
            {} as unknown as ChatContext,
            response as unknown as ChatResponseStream,
            token
          )
        )
        .to.equal(true);
    });

    it("call defaultHandler", async () => {
      const request = {
        prompt: "fakePrompt",
        command: "",
        references: [],
        location: 1,
        attempt: 0,
        enableCommandDetection: false,
      } as ChatRequest;

      const chatTelemetryDataMock = sandbox.createStubInstance(telemetry.ChatTelemetryData);
      const metaDataMock = { metadata: { command: undefined, requestId: undefined } };
      sandbox.stub(chatTelemetryDataMock, "properties").get(function getterFn() {
        return undefined;
      });
      sandbox.stub(chatTelemetryDataMock, "measurements").get(function getterFn() {
        return undefined;
      });
      chatTelemetryDataMock.chatMessages = [];
      sandbox
        .stub(telemetry.ChatTelemetryData, "createByParticipant")
        .returns(chatTelemetryDataMock);
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(util, "verbatimCopilotInteraction");
      const result = await handler.chatRequestHandler(
        request,
        {} as unknown as ChatContext,
        response as unknown as ChatResponseStream,
        token
      );

      chai.expect(result).to.deep.equal(metaDataMock);
    });

    it("call defaultHandler - error", async () => {
      const request = {
        prompt: "",
        command: "",
        references: [],
        location: 1,
        attempt: 0,
        enableCommandDetection: false,
      } as ChatRequest;

      const chatTelemetryDataMock = sandbox.createStubInstance(telemetry.ChatTelemetryData);
      const metaDataMock = { metadata: { command: undefined, requestId: undefined } };
      sandbox.stub(chatTelemetryDataMock, "properties").get(function getterFn() {
        return undefined;
      });
      sandbox.stub(chatTelemetryDataMock, "measurements").get(function getterFn() {
        return undefined;
      });
      chatTelemetryDataMock.chatMessages = [];
      sandbox
        .stub(telemetry.ChatTelemetryData, "createByParticipant")
        .returns(chatTelemetryDataMock);
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(util, "verbatimCopilotInteraction");
      await chai.expect(
        handler.chatRequestHandler(
          request,
          {} as unknown as ChatContext,
          response as unknown as ChatResponseStream,
          token
        )
      ).is.rejectedWith(`
Please specify a question when using this command.

Usage: @teams Ask questions about Teams Development"`);
    });
  });

  describe("chatExecuteCommandHandler()", () => {
    const sandbox = sinon.createSandbox();
    afterEach(async () => {
      sandbox.restore();
    });

    it("execute commands", async () => {
      const chatTelemetryDataMock = sandbox.createStubInstance(telemetry.ChatTelemetryData);
      sandbox.stub(chatTelemetryDataMock, "properties").get(function getterFn() {
        return undefined;
      });
      sandbox.stub(chatTelemetryDataMock, "measurements").get(function getterFn() {
        return undefined;
      });
      chatTelemetryDataMock.requestId = "fakeRequestId";
      sandbox.stub(telemetry.ChatTelemetryData, "get").returns(chatTelemetryDataMock);
      const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      const executeCommandStub = sandbox.stub(commands, "executeCommand");
      await handler.chatExecuteCommandHandler("fakeCommand", "fakeRequestId", ["fakeArgs"]);

      chai.expect(sendTelemetryEventStub.calledOnce).to.equal(true);
      chai.expect(executeCommandStub.calledOnce).to.equal(true);
    });

    it("execute commands with undefined chat telemetry data", async () => {
      sandbox.stub(telemetry.ChatTelemetryData, "get").returns(undefined);
      const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      const executeCommandStub = sandbox.stub(commands, "executeCommand");
      await handler.chatExecuteCommandHandler(CommandKey.OpenReadMe, "fakeRequestId", ["fakeArgs"]);

      chai.expect(sendTelemetryEventStub.called).to.equal(false);
      chai.expect(executeCommandStub.calledOnce).to.equal(true);
    });
  });

  describe("openUrlCommandHandler()", () => {
    const sandbox = sinon.createSandbox();
    afterEach(async () => {
      sandbox.restore();
    });

    it("open external", async () => {
      await openUrlCommandHandler("fakeUrl");
    });
  });

  describe("handleFeedback()", () => {
    const sandbox = sinon.createSandbox();
    afterEach(async () => {
      sandbox.restore();
    });

    it("handle feedback with undefined request id and command", async () => {
      const fakeFeedback: ChatResultFeedback = {
        result: {},
        kind: 1,
      };
      sandbox.stub(Correlator, "getId").returns("testCorrelationId");
      const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      handler.handleFeedback(fakeFeedback);

      chai.expect(sendTelemetryEventStub.calledOnce).to.equal(true);
      chai.expect(sendTelemetryEventStub.args[0]).to.deep.equal([
        TelemetryEvent.CopilotChatFeedback,
        {
          [TelemetryProperty.CopilotChatRequestId]: "",
          [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.CopilotChat,
          [TelemetryProperty.CopilotChatCommand]: "",
          [TelemetryProperty.CorrelationId]: "testCorrelationId",
        },
        {
          [TelemetryProperty.CopilotChatFeedbackHelpful]: 1,
        },
      ]);
    });

    it("handle feedback with request id and command", async () => {
      const fakeFeedback: ChatResultFeedback = {
        result: {
          metadata: {
            requestId: "testRequestId",
            command: "testCommand",
          },
        },
        kind: 0,
      };
      sandbox.stub(Correlator, "getId").returns("testCorrelationId");
      const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      handler.handleFeedback(fakeFeedback);

      chai.expect(sendTelemetryEventStub.calledOnce).to.equal(true);
      chai.expect(sendTelemetryEventStub.args[0]).to.deep.equal([
        TelemetryEvent.CopilotChatFeedback,
        {
          [TelemetryProperty.CopilotChatRequestId]: "testRequestId",
          [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.CopilotChat,
          [TelemetryProperty.CopilotChatCommand]: "testCommand",
          [TelemetryProperty.CorrelationId]: "testCorrelationId",
        },
        {
          [TelemetryProperty.CopilotChatFeedbackHelpful]: 0,
        },
      ]);
    });
  });
});

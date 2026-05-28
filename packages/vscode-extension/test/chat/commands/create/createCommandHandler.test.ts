import * as chai from "chai";
import chaiPromised from "chai-as-promised";
import * as sinon from "sinon";
import * as vscode from "vscode";
import * as createCommandHandler from "../../../../src/chat/commands/create/createCommandHandler";
import * as helper from "../../../../src/chat/commands/create/helper";
import { ProjectMetadata } from "../../../../src/chat/commands/create/types";
import * as telemetry from "../../../../src/chat/telemetry";
import * as util from "../../../../src/chat/utils";
import { ExtTelemetry } from "../../../../src/telemetry/extTelemetry";
import { CancellationToken } from "../../../mocks/vsc";

chai.use(chaiPromised);

describe("chat create command", () => {
  afterEach(() => {
    sinon.restore();
  });

  describe("createCommandHandler()", () => {
    const sandbox = sinon.createSandbox();

    afterEach(async () => {
      sandbox.restore();
    });

    it("returns default answer", async () => {
      const chatTelemetryDataMock = sandbox.createStubInstance(telemetry.ChatTelemetryData);
      sandbox.stub(chatTelemetryDataMock, "properties").get(function getterFn() {
        return undefined;
      });
      sandbox.stub(chatTelemetryDataMock, "measurements").get(function getterFn() {
        return undefined;
      });
      sandbox
        .stub(telemetry.ChatTelemetryData, "createByParticipant")
        .returns(chatTelemetryDataMock);
      const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      const response = {
        markdown: sandbox.stub(),
      };
      const token = new CancellationToken();
      await createCommandHandler.default(
        { prompt: "" } as unknown as vscode.ChatRequest,
        {} as unknown as vscode.ChatContext,
        response as unknown as vscode.ChatResponseStream,
        token
      );
      chai.assert.isTrue(
        response.markdown.calledOnceWith(
          "Use this command to provide description and other details about the Teams app that you want to build.\n\nE.g. @teams /create a Teams app that will notify my team about new GitHub pull requests.\n\n@teams /create I want to create a ToDo Teams app."
        )
      );
      chai.assert.isTrue(sendTelemetryEventStub.calledTwice);
    });

    it("returns no result answer", async () => {
      const chatTelemetryDataMock = sandbox.createStubInstance(telemetry.ChatTelemetryData);
      sandbox.stub(chatTelemetryDataMock, "properties").get(function getterFn() {
        return undefined;
      });
      sandbox.stub(chatTelemetryDataMock, "measurements").get(function getterFn() {
        return undefined;
      });
      sandbox
        .stub(telemetry.ChatTelemetryData, "createByParticipant")
        .returns(chatTelemetryDataMock);
      const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      const matchProjectStub = sandbox.stub(helper, "matchProject").resolves([]);

      const response = {
        markdown: sandbox.stub(),
      };
      const token = new CancellationToken();
      await createCommandHandler.default(
        { prompt: "test" } as unknown as vscode.ChatRequest,
        {} as unknown as vscode.ChatContext,
        response as unknown as vscode.ChatResponseStream,
        token
      );
      chai.assert.isTrue(
        response.markdown.calledOnceWith(
          "I cannot find any matching templates or samples. Refine your app description or explore other templates."
        )
      );
    });

    it("has exactly 1 matched sample", async () => {
      const chatTelemetryDataMock = sandbox.createStubInstance(telemetry.ChatTelemetryData);
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
      const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      const fakedSample = {
        id: "test-sample",
        type: "sample",
        platform: "Teams",
        name: "test sample",
        description: "test sample",
      } as ProjectMetadata;
      sandbox.stub(helper, "matchProject").resolves([fakedSample]);
      const showFileTreeStub = sandbox.stub(helper, "showFileTree");
      sandbox.stub(util, "verbatimCopilotInteraction");

      const response = {
        markdown: sandbox.stub(),
        button: sandbox.stub(),
      };
      const token = new CancellationToken();
      await createCommandHandler.default(
        { prompt: "test" } as unknown as vscode.ChatRequest,
        {} as unknown as vscode.ChatContext,
        response as unknown as vscode.ChatResponseStream,
        token
      );
      chai.assert.isTrue(showFileTreeStub.calledOnce);
    });

    it("has exactly 1 matched template", async () => {
      const chatTelemetryDataMock = sandbox.createStubInstance(telemetry.ChatTelemetryData);
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
      const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      const fakedSample = {
        id: "test-template",
        type: "template",
        platform: "Teams",
        name: "test template",
        description: "test template",
      } as ProjectMetadata;
      sandbox.stub(helper, "matchProject").resolves([fakedSample]);
      const showFileTreeStub = sandbox.stub(helper, "showFileTree");
      sandbox.stub(util, "verbatimCopilotInteraction");

      const response = {
        markdown: sandbox.stub(),
        button: sandbox.stub(),
      };
      const token = new CancellationToken();
      await createCommandHandler.default(
        { prompt: "test" } as unknown as vscode.ChatRequest,
        {} as unknown as vscode.ChatContext,
        response as unknown as vscode.ChatResponseStream,
        token
      );
      chai.assert.isTrue(showFileTreeStub.notCalled);
    });

    it("has multiple matched results", async () => {
      const chatTelemetryDataMock = sandbox.createStubInstance(telemetry.ChatTelemetryData);
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
      const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      const fakedSamples = [
        {
          id: "test-sample",
          type: "sample",
          platform: "Teams",
          name: "test sample",
          description: "test sample",
        },
        {
          id: "test-sample",
          type: "template",
          platform: "Teams",
          name: "test sample",
          description: "test sample",
        },
      ] as ProjectMetadata[];
      sandbox.stub(helper, "matchProject").resolves(fakedSamples);
      const showFileTreeStub = sandbox.stub(helper, "showFileTree");
      sandbox.stub(util, "verbatimCopilotInteraction");

      const response = {
        markdown: sandbox.stub(),
        button: sandbox.stub(),
      };
      const token = new CancellationToken();
      await createCommandHandler.default(
        { prompt: "test" } as unknown as vscode.ChatRequest,
        {} as unknown as vscode.ChatContext,
        response as unknown as vscode.ChatResponseStream,
        token
      );
      chai.assert.isTrue(showFileTreeStub.notCalled);
      chai.assert.isTrue(response.markdown.calledThrice);
      chai.assert.isTrue(response.button.calledTwice);
    });

    it("has >5 matched results", async () => {
      const chatTelemetryDataMock = sandbox.createStubInstance(telemetry.ChatTelemetryData);
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
      const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      const fakedSamples = [
        {
          id: "test-sample",
          type: "sample",
          platform: "Teams",
          name: "test sample",
          description: "test sample",
        },
        {
          id: "test-sample",
          type: "template",
          platform: "Teams",
          name: "test sample",
          description: "test sample",
        },
        {
          id: "test-sample",
          type: "sample",
          platform: "Teams",
          name: "test sample",
          description: "test sample",
        },
        {
          id: "test-sample",
          type: "template",
          platform: "Teams",
          name: "test sample",
          description: "test sample",
        },
        {
          id: "test-sample",
          type: "sample",
          platform: "Teams",
          name: "test sample",
          description: "test sample",
        },
        {
          id: "test-sample",
          type: "template",
          platform: "Teams",
          name: "test sample",
          description: "test sample",
        },
      ] as ProjectMetadata[];
      sandbox.stub(helper, "matchProject").resolves(fakedSamples);
      const showFileTreeStub = sandbox.stub(helper, "showFileTree");
      sandbox.stub(util, "verbatimCopilotInteraction");

      const response = {
        markdown: sandbox.stub(),
        button: sandbox.stub(),
      };
      const token = new CancellationToken();
      await createCommandHandler.default(
        { prompt: "test" } as unknown as vscode.ChatRequest,
        {} as unknown as vscode.ChatContext,
        response as unknown as vscode.ChatResponseStream,
        token
      );
      chai.assert.isTrue(showFileTreeStub.notCalled);
      chai.assert.isTrue(
        response.markdown.calledOnceWith(
          "Your app description is too generic. To find relevant templates or samples, give specific details of your app's capabilities or technologies.\n\nE.g. Instead of saying 'create a bot', you could specify 'create a bot template' or 'create a notification bot that sends user the stock updates'."
        )
      );
    });
  });
});

import * as chai from "chai";
import * as sinon from "sinon";
import chaiPromised from "chai-as-promised";
import * as vscode from "vscode";
import * as util from "../../../../src/officeChat/utils";
import * as helper from "../../../../src/officeChat/commands/create/helper";
import * as generatecodeCommandHandler from "../../../../src/officeChat/commands/generatecode/generatecodeCommandHandler";
import { ExtTelemetry } from "../../../../src/telemetry/extTelemetry";
import { CancellationToken } from "../../../mocks/vsc";
import { Planner } from "../../../../src/officeChat/common/planner";
import { OfficeChatTelemetryData } from "../../../../src/officeChat/telemetry";

chai.use(chaiPromised);

describe("File: generatecodeCommandHandler", () => {
  const sandbox = sinon.createSandbox();
  let sendTelemetryEventStub: any;
  let officeChatTelemetryDataMock: any;
  beforeEach(() => {
    officeChatTelemetryDataMock = sandbox.createStubInstance(OfficeChatTelemetryData);
    sandbox.stub(officeChatTelemetryDataMock, "properties").get(function getterFn() {
      return undefined;
    });
    sandbox.stub(officeChatTelemetryDataMock, "measurements").get(function getterFn() {
      return undefined;
    });
    officeChatTelemetryDataMock.chatMessages = [];
    sandbox
      .stub(OfficeChatTelemetryData, "createByParticipant")
      .returns(officeChatTelemetryDataMock);
    sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
  });

  afterEach(() => {
    sandbox.restore();
    process.env.NODE_ENV = undefined;
  });

  it("input prompt is empty", async () => {
    const response = {
      markdown: sandbox.stub(),
    };
    const token = new CancellationToken();
    await generatecodeCommandHandler.default(
      { prompt: "" } as unknown as vscode.ChatRequest,
      {} as unknown as vscode.ChatContext,
      response as unknown as vscode.ChatResponseStream,
      token
    );
    chai.assert.isTrue(
      response.markdown.calledOnceWith(
        "Use this command to provide description and other details about the code snippets you want to try.\n\nE.g. @office /generatecode @office /generatecode create a chart based on the selected range in Excel.\n\n@office /generatecode @office /generatecode insert a content control in a Word document."
      )
    );
    chai.assert.isTrue(sendTelemetryEventStub.calledTwice);
  });

  it("input prompt is empty in dev env", async () => {
    process.env.NODE_ENV = "development";
    const response = {
      markdown: sandbox.stub(),
    };
    const token = new CancellationToken();
    await generatecodeCommandHandler.default(
      { prompt: "" } as unknown as vscode.ChatRequest,
      {} as unknown as vscode.ChatContext,
      response as unknown as vscode.ChatResponseStream,
      token
    );
    chai.assert.isTrue(
      response.markdown.calledOnceWith(
        "Use this command to provide description and other details about the code snippets you want to try.\n\nE.g. @office /generatecode @office /generatecode create a chart based on the selected range in Excel.\n\n@office /generatecode @office /generatecode insert a content control in a Word document."
      )
    );
    chai.assert.isTrue(sendTelemetryEventStub.calledTwice);
  });

  it("input prompt is harmful", async () => {
    const response = {
      markdown: sandbox.stub(),
    };
    const isInputHarmfulStub = sandbox.stub(util, "isInputHarmful").resolves(true);
    const token = new CancellationToken();
    await generatecodeCommandHandler.default(
      { prompt: "test" } as unknown as vscode.ChatRequest,
      {} as unknown as vscode.ChatContext,
      response as unknown as vscode.ChatResponseStream,
      token
    );
    chai.assert.isTrue(isInputHarmfulStub.calledOnce);
    chai.assert.isTrue(response.markdown.calledOnceWith("Sorry, I can't assist with that."));
  });

  it("should call the planner to process the request", async () => {
    const processRequestStub = sandbox.stub(Planner.getInstance(), "processRequest");
    const response = {
      markdown: sandbox.stub(),
    };
    const token = new CancellationToken();
    sandbox.stub(util, "isInputHarmful").resolves(false);
    sandbox.stub(helper, "matchOfficeProject").resolves(undefined);
    await generatecodeCommandHandler.default(
      { prompt: "test" } as unknown as vscode.ChatRequest,
      {} as unknown as vscode.ChatContext,
      response as unknown as vscode.ChatResponseStream,
      token
    );
    chai.assert.isTrue(processRequestStub.calledOnce);
  });
});

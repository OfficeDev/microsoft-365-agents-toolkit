import * as chai from "chai";
import * as sinon from "sinon";
import chaiPromised from "chai-as-promised";
import * as vscode from "vscode";
import * as officeCreateCommandHandler from "../../../../src/officeChat/commands/create/officeCreateCommandHandler";
import * as officeChatUtil from "../../../../src/officeChat/utils";
import * as helper from "../../../../src/officeChat/commands/create/helper";
import * as chatUtil from "../../../../src/chat/utils";
import { ExtTelemetry } from "../../../../src/telemetry/extTelemetry";
import { CancellationToken } from "../../../mocks/vsc";
import { ProjectMetadata } from "../../../../src/chat/commands/create/types";
import { Planner } from "../../../../src/officeChat/common/planner";
import { OfficeChatTelemetryData } from "../../../../src/officeChat/telemetry";
import { OfficeProjectInfo } from "../../../../src/officeChat/types";

chai.use(chaiPromised);

describe("File: officeCreateCommandHandler", () => {
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
  });

  it("input prompt is empty", async () => {
    const response = {
      markdown: sandbox.stub(),
    };
    const token = new CancellationToken();
    await officeCreateCommandHandler.default(
      { prompt: "" } as unknown as vscode.ChatRequest,
      {} as unknown as vscode.ChatContext,
      response as unknown as vscode.ChatResponseStream,
      token
    );
    chai.assert.isTrue(
      response.markdown.calledOnceWith(
        "Use this command to provide description and other details about the Office Add-ins that you want to build.\n\nE.g. @office /create an Excel hello world Add-in.\n\n@office /create a Word Add-in that inserts comments."
      )
    );
    chai.assert.isTrue(sendTelemetryEventStub.calledTwice);
  });

  it("input prompt is harmful", async () => {
    const response = {
      markdown: sandbox.stub(),
    };
    const isInputHarmfulStub = sandbox.stub(officeChatUtil, "isInputHarmful").resolves(true);
    const token = new CancellationToken();
    await officeCreateCommandHandler.default(
      { prompt: "test" } as unknown as vscode.ChatRequest,
      {} as unknown as vscode.ChatContext,
      response as unknown as vscode.ChatResponseStream,
      token
    );
    chai.assert.isTrue(isInputHarmfulStub.calledOnce);
    chai.assert.isTrue(response.markdown.calledOnceWith("Sorry, I can't assist with that."));
  });

  it("has 1 matched sample", async () => {
    const fakedSample = {
      id: "test-sample",
      type: "sample",
      platform: "WXP",
      name: "test sample",
      description: "test sample",
    } as ProjectMetadata;
    sandbox.stub(officeChatUtil, "isInputHarmful").resolves(false);
    sandbox.stub(helper, "matchOfficeProject").resolves(fakedSample);
    const mockOfficeProjectInfo: OfficeProjectInfo = {
      path: "",
      host: "",
    };
    const showOfficeSampleFileTreeStub = sandbox
      .stub(helper, "showOfficeSampleFileTree")
      .resolves(mockOfficeProjectInfo);
    sandbox.stub(chatUtil, "verbatimCopilotInteraction");
    const response = {
      markdown: sandbox.stub(),
      button: sandbox.stub(),
    };
    const token = new CancellationToken();
    await officeCreateCommandHandler.default(
      { prompt: "test" } as unknown as vscode.ChatRequest,
      {} as unknown as vscode.ChatContext,
      response as unknown as vscode.ChatResponseStream,
      token
    );
    chai.assert.isTrue(showOfficeSampleFileTreeStub.calledOnce);
    chai.assert.isTrue(response.button.calledOnce);
  });

  it("has 1 matched template", async () => {
    const fakedTemplate = {
      id: "test-id",
      type: "template",
      platform: "WXP",
      name: "test template",
      description: "test template",
    } as ProjectMetadata;
    sandbox.stub(officeChatUtil, "isInputHarmful").resolves(false);
    sandbox.stub(helper, "matchOfficeProject").resolves(fakedTemplate);
    const showOfficeSampleFileTreeStub = sandbox.stub(helper, "showOfficeTemplateFileTree");
    sandbox.stub(chatUtil, "verbatimCopilotInteraction");
    const response = {
      markdown: sandbox.stub(),
      button: sandbox.stub(),
    };
    const token = new CancellationToken();
    await officeCreateCommandHandler.default(
      { prompt: "test" } as unknown as vscode.ChatRequest,
      {} as unknown as vscode.ChatContext,
      response as unknown as vscode.ChatResponseStream,
      token
    );
    chai.assert.isTrue(showOfficeSampleFileTreeStub.calledOnce);
    chai.assert.isTrue(response.button.calledOnce);
  });

  it("should call the planner to process the request", async () => {
    const processRequestStub = sandbox.stub(Planner.getInstance(), "processRequest");
    const response = {
      markdown: sandbox.stub(),
    };
    const token = new CancellationToken();
    sandbox.stub(officeChatUtil, "isInputHarmful").resolves(false);
    sandbox.stub(helper, "matchOfficeProject").resolves(undefined);
    await officeCreateCommandHandler.default(
      { prompt: "test" } as unknown as vscode.ChatRequest,
      {} as unknown as vscode.ChatContext,
      response as unknown as vscode.ChatResponseStream,
      token
    );
    chai.assert.isTrue(processRequestStub.calledOnce);
  });
});

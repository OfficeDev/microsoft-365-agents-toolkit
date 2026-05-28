import * as chai from "chai";
import * as sinon from "sinon";
import chaiPromised from "chai-as-promised";
import * as vscode from "vscode";
import fs from "fs-extra";
import path from "path";
import os from "os";
import * as handler from "../../src/officeChat/handlers";
import * as util from "../../src/chat/utils";
import * as localizeUtils from "../../src/utils/localizeUtils";
import * as officeCreateCommandHandler from "../../src/officeChat/commands/create/officeCreateCommandHandler";
import * as generatecodeCommandHandler from "../../src/officeChat/commands/generatecode/generatecodeCommandHandler";
import * as officeNextStepCommandHandler from "../../src/officeChat/commands/nextStep/officeNextstepCommandHandler";
import * as workspaceUtils from "../../src/utils/workspaceUtils";
import { URI } from "../mocks/vsc/uri";
import { OfficeChatCommand } from "../../src/officeChat/consts";
import { CancellationToken } from "../mocks/vsc";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetryTriggerFrom,
} from "../../src/telemetry/extTelemetryEvents";
import { Correlator } from "@microsoft/teamsfx-core";
import { ConstantString } from "@microsoft/teamsfx-core/build/common/constants";
import { OfficeChatTelemetryData } from "../../src/officeChat/telemetry";

chai.use(chaiPromised);

describe("File: officeChat/handlers.ts", () => {
  describe("Method: officeChatRequestHandler", () => {
    const sandbox = sinon.createSandbox();
    const response = {
      markdown: sandbox.stub(),
      button: sandbox.stub(),
    };
    const token = new CancellationToken();
    afterEach(() => {
      sandbox.restore();
    });

    it("call officeCreateCommandHandler", async () => {
      const request = {
        prompt: "test",
        command: OfficeChatCommand.Create,
        references: [],
        location: 1,
        attempt: 0,
        enableCommandDetection: false,
      } as vscode.ChatRequest;
      const officeCreateCommandHandlerStub = sandbox.stub(officeCreateCommandHandler, "default");
      handler.officeChatRequestHandler(
        request,
        {} as unknown as vscode.ChatContext,
        response as unknown as vscode.ChatResponseStream,
        token
      );
      chai.expect(officeCreateCommandHandlerStub.calledOnce);
    });

    it("call generatecodeCommandHandler", async () => {
      const request = {
        prompt: "test",
        command: OfficeChatCommand.GenerateCode,
        references: [],
        location: 1,
        attempt: 0,
        enableCommandDetection: false,
      } as vscode.ChatRequest;
      const generatecodeCommandHandlerStub = sandbox.stub(generatecodeCommandHandler, "default");
      handler.officeChatRequestHandler(
        request,
        {} as unknown as vscode.ChatContext,
        response as unknown as vscode.ChatResponseStream,
        token
      );
      chai.expect(generatecodeCommandHandlerStub.calledOnce);
    });

    it("call officeNextStepCommandHandler", async () => {
      const request = {
        prompt: "test",
        command: OfficeChatCommand.NextStep,
        references: [],
        location: 1,
        attempt: 0,
        enableCommandDetection: false,
      } as vscode.ChatRequest;
      const officeNextStepCommandHandlerStub = sandbox.stub(
        officeNextStepCommandHandler,
        "default"
      );
      handler.officeChatRequestHandler(
        request,
        {} as unknown as vscode.ChatContext,
        response as unknown as vscode.ChatResponseStream,
        token
      );
      chai.expect(officeNextStepCommandHandlerStub.calledOnce);
    });

    it("call officeDefaultHandler", async () => {
      const request = {
        prompt: "test",
        command: "",
        references: [],
        location: 1,
        attempt: 0,
        enableCommandDetection: false,
      } as vscode.ChatRequest;
      const officeChatTelemetryDataMock = sandbox.createStubInstance(OfficeChatTelemetryData);
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
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      const verbatimCopilotInteractionStub = sandbox.stub(util, "verbatimCopilotInteraction");
      await handler.officeChatRequestHandler(
        request,
        {} as unknown as vscode.ChatContext,
        response as unknown as vscode.ChatResponseStream,
        token
      );
      chai.expect(verbatimCopilotInteractionStub.calledOnce);
    });

    it("call officeDefaultHandler - error", async () => {
      const request = {
        prompt: "",
        command: "",
        references: [],
        location: 1,
        attempt: 0,
        enableCommandDetection: false,
      } as vscode.ChatRequest;
      const officeChatTelemetryDataMock = sandbox.createStubInstance(OfficeChatTelemetryData);
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
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(util, "verbatimCopilotInteraction");
      await chai.expect(
        handler.officeChatRequestHandler(
          request,
          {} as unknown as vscode.ChatContext,
          response as unknown as vscode.ChatResponseStream,
          token
        )
      ).is.rejectedWith(`
Please specify a question when using this command.

Usage: @office Ask questions about Office Add-ins development.`);
    });
  });

  describe("method: chatCreateOfficeProjectCommandHandler", () => {
    const sandbox = sinon.createSandbox();
    const defaultFolder = path.join(os.homedir(), ConstantString.RootFolder);
    afterEach(async () => {
      sandbox.restore();
    });

    it("choose no folder", async () => {
      const fsCopyStub = sandbox.stub(fs, "copy");
      const showQuickPickStub = sandbox
        .stub(vscode.window, "showQuickPick")
        .returns(Promise.resolve(undefined));
      const result = await handler.chatCreateOfficeProjectCommandHandler(
        "fakeFolder",
        "fakeId",
        "fakeMatchResultInfo",
        "fakeAppId"
      );

      chai.expect(result).to.equal(undefined);
      chai.expect(showQuickPickStub.calledOnce).to.equal(true);
      chai.expect(fsCopyStub.called).to.equal(false);
    });

    it("choose default folder", async () => {
      const showQuickPickStub = sandbox.stub(vscode.window, "showQuickPick").returns(
        Promise.resolve({
          label: "Default folder",
          description: defaultFolder,
        }) as unknown as Promise<vscode.QuickPickItem>
      );
      const fsCopyStub = sandbox.stub(fs, "copy");
      const showOpenDialogStub = sandbox.stub(vscode.window, "showOpenDialog");
      const openOfficeDevFolderStub = sandbox.stub(workspaceUtils, "openOfficeDevFolder");
      sandbox.stub(localizeUtils, "localize").returns("Default folder");
      sandbox.stub(fs, "pathExistsSync").returns(false);
      await handler.chatCreateOfficeProjectCommandHandler(
        "fakeFolder",
        "fakeId",
        "fakeMatchResultInfo",
        "fakeAppId"
      );

      chai.expect(showQuickPickStub.calledOnce).to.equal(true);
      chai.expect(showOpenDialogStub.called).to.equal(false);
      chai
        .expect(fsCopyStub.args[0])
        .to.deep.equal(["fakeFolder", path.join(defaultFolder, "fakeAppId")]);
      chai.expect(fsCopyStub.calledOnce).to.equal(true);
      chai.expect(openOfficeDevFolderStub.calledOnce).to.equal(true);
    });

    it("choose default folder but have naming conflicts", async () => {
      const showQuickPickStub = sandbox.stub(vscode.window, "showQuickPick").returns(
        Promise.resolve({
          label: "Default folder",
          description: defaultFolder,
        }) as unknown as Promise<vscode.QuickPickItem>
      );
      const fsCopyStub = sandbox.stub(fs, "copy");
      const showOpenDialogStub = sandbox.stub(vscode.window, "showOpenDialog");
      sandbox.stub(localizeUtils, "localize").returns("Default folder");
      const pathExistsSyncStub = sandbox.stub(fs, "pathExistsSync");
      pathExistsSyncStub.withArgs(path.join(defaultFolder, "fakeAppId")).returns(true);
      sandbox
        .stub(fs, "readdirSync")
        .returns([path.join(defaultFolder, "fakeAppId") as any as fs.Dirent]);
      await handler.chatCreateOfficeProjectCommandHandler(
        "fakeFolder",
        "fakeId",
        "fakeMatchResultInfo",
        "fakeAppId"
      );

      chai.expect(showQuickPickStub.calledOnce).to.equal(true);
      chai.expect(showOpenDialogStub.called).to.equal(false);
      chai
        .expect(fsCopyStub.args[0])
        .to.deep.equal(["fakeFolder", path.join(defaultFolder, "fakeAppId_1")]);
      chai.expect(fsCopyStub.calledOnce).to.equal(true);
    });

    it("choose to browse and select no folder", async () => {
      const showQuickPickStub = sandbox.stub(vscode.window, "showQuickPick").returns(
        Promise.resolve({
          label: "Browse...",
        }) as unknown as Promise<vscode.QuickPickItem>
      );
      const fsCopyStub = sandbox.stub(fs, "copy");
      const showOpenDialogStub = sandbox
        .stub(vscode.window, "showOpenDialog")
        .returns(Promise.resolve(undefined));
      sandbox.stub(localizeUtils, "localize").returns("Default folder");
      await handler.chatCreateOfficeProjectCommandHandler(
        "fakeFolder",
        "fakeId",
        "fakeMatchResultInfo",
        "fakeAppId"
      );

      chai.expect(showQuickPickStub.calledOnce).to.equal(true);
      chai.expect(showOpenDialogStub.calledOnce).to.equal(true);
      chai.expect(fsCopyStub.called).to.equal(false);
    });

    it("choose to browse and select custom folder", async () => {
      const showQuickPickStub = sandbox.stub(vscode.window, "showQuickPick").resolves({
        label: "Browse...",
      } as unknown as vscode.QuickPickItem);
      const fsCopyStub = sandbox.stub(fs, "copy");
      const customFolderPath = "customFolderPath";
      const customFolder: URI[] = [URI.file(customFolderPath)];
      const showOpenDialogStub = sandbox
        .stub(vscode.window, "showOpenDialog")
        .resolves(customFolder);
      sandbox.stub(fs, "pathExistsSync").returns(false);
      sandbox.stub(localizeUtils, "localize").returns("Default folder");
      sandbox.stub(fs, "ensureDirSync");
      await handler.chatCreateOfficeProjectCommandHandler(
        "fakeFolder",
        "fakeId",
        "fakeMatchResultInfo",
        "fakeAppId"
      );

      chai.expect(showQuickPickStub.calledOnce).to.equal(true);
      chai.expect(showOpenDialogStub.calledOnce).to.equal(true);
      chai.expect(fsCopyStub.calledOnce).to.equal(true);
      chai.expect(fsCopyStub.args[0][0]).to.equal("fakeFolder");
      chai.expect(path.basename(fsCopyStub.args[0][1])).to.equal("fakeAppId");
    });

    it("copy files error", async () => {
      const copyError = new Error("fakeError");
      const showQuickPickStub = sandbox.stub(vscode.window, "showQuickPick").returns(
        Promise.resolve({
          label: "Default folder",
          description: defaultFolder,
        }) as unknown as Promise<vscode.QuickPickItem>
      );
      const fsCopyStub = sandbox.stub(fs, "copy").throwsException(copyError);
      const showOpenDialogStub = sandbox.stub(vscode.window, "showOpenDialog");
      const showErrorMessageStub = sandbox.stub(vscode.window, "showErrorMessage");
      const consoleLogStub = sandbox.stub(console, "error");
      sandbox.stub(localizeUtils, "localize").callsFake((key: string) => {
        if (key === "teamstoolkit.chatParticipants.officeAddIn.create.failToCreate")
          return "Fail to Create";
        else return "Default folder";
      });
      await handler.chatCreateOfficeProjectCommandHandler(
        "fakeFolder",
        "fakeId",
        "fakeMatchResultInfo",
        "fakeAppId"
      );

      chai.expect(showQuickPickStub.calledOnce).to.equal(true);
      chai.expect(showOpenDialogStub.called).to.equal(false);
      chai.expect(fsCopyStub.calledOnce).to.equal(true);
      chai.expect(consoleLogStub.args[0][0]).to.equal("Error copying files:");
      chai.expect(consoleLogStub.args[0][1]).to.deep.equal(copyError);
      chai.expect(consoleLogStub.calledOnce).to.equal(true);
      chai.expect(showErrorMessageStub.args[0]).to.deep.equal(["Fail to Create"]);
      chai.expect(showErrorMessageStub.calledOnce).to.equal(true);
    });
  });

  describe("Method: handleOfficeFeedback", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("handle feedback with undefined request id and command", async () => {
      const fakedFeedback: vscode.ChatResultFeedback = {
        result: {},
        kind: 1,
      };
      sandbox.stub(Correlator, "getId").returns("testCorrelationId");
      const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      handler.handleOfficeFeedback(fakedFeedback);

      chai.expect(sendTelemetryEventStub.calledOnce).to.equal(true);
      chai.expect(sendTelemetryEventStub.args[0]).to.deep.equal([
        TelemetryEvent.CopilotChatFeedback,
        {
          [TelemetryProperty.CopilotChatRequestId]: "",
          [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.CopilotChat,
          [TelemetryProperty.CopilotChatCommand]: "",
          [TelemetryProperty.CorrelationId]: "testCorrelationId",
          [TelemetryProperty.HostType]: "",
          [TelemetryProperty.CopilotChatRelatedSampleName]: "",
        },
        {
          [TelemetryProperty.CopilotChatFeedbackHelpful]: 1,
        },
      ]);
    });

    it("handle feedback with request id and command", async () => {
      const fakeFeedback: vscode.ChatResultFeedback = {
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
      handler.handleOfficeFeedback(fakeFeedback);

      chai.expect(sendTelemetryEventStub.calledOnce).to.equal(true);
      chai.expect(sendTelemetryEventStub.args[0]).to.deep.equal([
        TelemetryEvent.CopilotChatFeedback,
        {
          [TelemetryProperty.CopilotChatRequestId]: "testRequestId",
          [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.CopilotChat,
          [TelemetryProperty.CopilotChatCommand]: "testCommand",
          [TelemetryProperty.CorrelationId]: "testCorrelationId",
          [TelemetryProperty.HostType]: "",
          [TelemetryProperty.CopilotChatRelatedSampleName]: "",
        },
        {
          [TelemetryProperty.CopilotChatFeedbackHelpful]: 0,
        },
      ]);
    });
  });
});

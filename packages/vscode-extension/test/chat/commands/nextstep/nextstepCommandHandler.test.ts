import * as chai from "chai";
import chaiPromised from "chai-as-promised";
import * as sinon from "sinon";
import * as vscode from "vscode";
import * as nextstepCommandHandler from "../../../../src/chat/commands/nextstep/nextstepCommandHandler";
import * as telemetry from "../../../../src/chat/telemetry";
import { ExtTelemetry } from "../../../../src/telemetry/extTelemetry";
import { CancellationToken } from "../../../mocks/vsc";
import * as globalVariables from "../../../../src/globalVariables";
import * as core from "@microsoft/teamsfx-core";
import * as status from "../../../../src/chat/commands/nextstep/status";
import { NextStep, WholeStatus } from "../../../../src/chat/commands/nextstep/types";
import * as steps from "../../../../src/chat/commands/nextstep/steps";
import { TeamsFollowupProvider } from "../../../../src/chat/followupProvider";
import * as util from "../../../../src/chat/utils";
import { CHAT_EXECUTE_COMMAND_ID, CHAT_OPENURL_COMMAND_ID } from "../../../../src/chat/consts";

chai.use(chaiPromised);

describe("chat nextstep handler", () => {
  afterEach(() => {
    sinon.restore();
  });

  describe("nextstepCommandHandler()", () => {
    const sandbox = sinon.createSandbox();

    afterEach(async () => {
      sandbox.restore();
    });

    it("prompt is unempty", async () => {
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
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      const response = {
        markdown: sandbox.stub(),
      };
      const token = new CancellationToken();
      await nextstepCommandHandler.default(
        {
          prompt: "123123",
        } as vscode.ChatRequest,
        {} as vscode.ChatContext,
        response as unknown as vscode.ChatResponseStream,
        token
      );
      chai.assert.isTrue(
        response.markdown.calledOnceWith(
          `This command provides guidance on your next steps based on your workspace.\n\nE.g. If you're unsure what to do after creating a project, simply ask Copilot by using @teams /nextstep.`
        )
      );
    });

    it("prompt empty - no workspace", async () => {
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
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      sandbox.stub(globalVariables, "workspaceUri").returns(undefined);
      sandbox.stub(core, "isValidProject").returns(false);
      sandbox.stub(status, "getWholeStatus").resolves({} as WholeStatus);
      sandbox.stub(steps, "allSteps").returns([
        {
          title: "selected - no workspace",
          description: (status) => "description: selected - no workspace",
          followUps: [],
          commands: [],
          condition: (status) => true,
          priority: 1,
        } as NextStep,
        {
          title: "selected - no workspace 2",
          description: (status) => "description: selected - no workspace 2",
          followUps: [],
          commands: [],
          condition: (status) => true,
          priority: 0,
        } as NextStep,
        {
          title: "not selected - no workspace",
          description: (status) => "description: not selected - no workspace",
          followUps: [],
          commands: [],
          condition: (status) => false,
          priority: 2,
        } as NextStep,
      ]);
      const getCopilotResponseAsStringStub = sandbox
        .stub(util, "getCopilotResponseAsString")
        .resolves("");
      const followupProviderStub = sandbox.stub(TeamsFollowupProvider.prototype, "addFollowups");

      const response = {
        markdown: sandbox.stub(),
      };
      const token = new CancellationToken();
      await nextstepCommandHandler.default(
        {} as vscode.ChatRequest,
        {} as vscode.ChatContext,
        response as unknown as vscode.ChatResponseStream,
        token
      );
      chai.assert.isTrue(getCopilotResponseAsStringStub.calledTwice);
      chai.assert.equal(response.markdown.callCount, 3);
      chai.assert.isTrue(followupProviderStub.calledOnce);
    });

    it("prompt empty - app opened", async () => {
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
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      sandbox.stub(globalVariables, "workspaceUri").returns(vscode.Uri.parse("test-workspace"));
      sandbox.stub(core, "isValidProject").returns(true);
      sandbox.stub(status, "getWholeStatus").resolves({} as WholeStatus);
      sandbox.stub(steps, "allSteps").returns([
        {
          title: "selected - app opened",
          description: "description: selected - app opened",
          followUps: [],
          docLink: "docLink",
          commands: [
            {
              command: CHAT_EXECUTE_COMMAND_ID,
              title: "title",
              arguments: ["command-name"],
            },
            {
              command: CHAT_OPENURL_COMMAND_ID,
              title: "title",
              arguments: ["url"],
            },
          ],
          condition: (status) => true,
          priority: 1,
        } as NextStep,
      ]);
      const getCopilotResponseAsStringStub = sandbox
        .stub(util, "getCopilotResponseAsString")
        .resolves("");
      const followupProviderStub = sandbox.stub(TeamsFollowupProvider.prototype, "addFollowups");

      const response = {
        markdown: sandbox.stub(),
        button: sandbox.stub(),
      };
      const token = new CancellationToken();
      await nextstepCommandHandler.default(
        {} as vscode.ChatRequest,
        {} as vscode.ChatContext,
        response as unknown as vscode.ChatResponseStream,
        token
      );
      chai.assert.isTrue(getCopilotResponseAsStringStub.calledOnce);
      chai.assert.isTrue(response.markdown.calledOnce);
      chai.assert.isTrue(response.button.calledTwice);
      chai.assert.isTrue(followupProviderStub.calledOnce);
    });
  });
});

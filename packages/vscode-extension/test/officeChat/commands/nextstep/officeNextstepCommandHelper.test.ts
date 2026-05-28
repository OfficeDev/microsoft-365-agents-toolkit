import * as chai from "chai";
import * as sinon from "sinon";
import officeNextStepCommandHandler from "../../../../src/officeChat/commands/nextStep/officeNextstepCommandHandler";
import { ExtTelemetry } from "../../../../src/telemetry/extTelemetry";
import { CancellationToken } from "../../../mocks/vsc";
import * as vscode from "vscode";
import * as globalVariables from "../../../../src/globalVariables";
import * as core from "@microsoft/teamsfx-core";
import { NextStep } from "../../../../src/chat/commands/nextstep/types";
import * as status from "../../../../src/officeChat/commands/nextStep/status";
import { TeamsFollowupProvider } from "../../../../src/chat/followupProvider";
import * as util from "../../../../src/chat/utils";
import * as officeSteps from "../../../../src/officeChat/commands/nextStep/officeSteps";
import { CHAT_EXECUTE_COMMAND_ID, CHAT_OPENURL_COMMAND_ID } from "../../../../src/chat/consts";
import { OfficeWholeStatus } from "../../../../src/officeChat/commands/nextStep/types";
import { OfficeChatTelemetryData } from "../../../../src/officeChat/telemetry";

afterEach(() => {
  // Restore the default sandbox here
  sinon.restore();
});

describe("office steps: officeNextStepCommandHandler", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    const chatTelemetryDataMock = sandbox.createStubInstance(OfficeChatTelemetryData);
    sandbox.stub(chatTelemetryDataMock, "properties").get(function getterFn() {
      return undefined;
    });
    sandbox.stub(chatTelemetryDataMock, "measurements").get(function getterFn() {
      return undefined;
    });
    chatTelemetryDataMock.chatMessages = [];
    chatTelemetryDataMock.responseChatMessages = [];
    sandbox.stub(OfficeChatTelemetryData, "createByParticipant").returns(chatTelemetryDataMock);
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("prompt is unempty", async () => {
    const response = {
      markdown: sandbox.stub(),
    };
    const token = new CancellationToken();

    await officeNextStepCommandHandler(
      {
        prompt: "123123",
      } as vscode.ChatRequest,
      {} as vscode.ChatContext,
      response as unknown as vscode.ChatResponseStream,
      token
    );
    chai.assert.isTrue(
      response.markdown.calledOnceWith(
        "This `/nextstep` command provides guidance on your next steps based on your workspace.\n\nE.g. To use this command, simply ask Copilot by using `@office /nextstep`."
      )
    );
  });

  it("prompt empty - no workspace", async () => {
    sandbox.stub(globalVariables, "workspaceUri").returns(undefined);
    sandbox.stub(core, "isValidOfficeAddInProject").returns(false);
    sandbox.stub(status, "getWholeStatus").resolves({} as OfficeWholeStatus);
    sandbox.stub(officeSteps, "officeSteps").returns([
      {
        title: "selected - no workspace",
        description: "description: selected - no workspace",
        followUps: [],
        commands: [],
        condition: (status) => true,
        priority: 1,
      } as NextStep,
    ]);
    const followupProviderStub = sandbox.stub(TeamsFollowupProvider.prototype, "addFollowups");

    const response = {
      markdown: sandbox.stub(),
    };
    const token = new CancellationToken();
    await officeNextStepCommandHandler(
      {} as vscode.ChatRequest,
      {} as vscode.ChatContext,
      response as unknown as vscode.ChatResponseStream,
      token
    );
    chai.assert.equal(response.markdown.callCount, 1);
    chai.assert.isTrue(followupProviderStub.calledOnce);
  });

  it("prompt empty - app opened", async () => {
    sandbox.stub(globalVariables, "workspaceUri").returns(undefined);
    sandbox.stub(core, "isValidOfficeAddInProject").returns(true);
    sandbox.stub(status, "getWholeStatus").resolves({} as OfficeWholeStatus);
    sandbox.stub(officeSteps, "officeSteps").returns([
      {
        title: "selected - app opened",
        description: () => "description: selected - app opened",
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
      {
        title: "selected 2 - app opened",
        description: () => "description: selected 2 - app opened",
        followUps: [],
        docLink: "docLink",
        commands: [
          {
            command: CHAT_EXECUTE_COMMAND_ID,
            title: "title",
            arguments: ["command-name"],
          },
        ],
        condition: (status) => true,
        priority: 1,
      } as NextStep,
    ]);
    const followupProviderStub = sandbox.stub(TeamsFollowupProvider.prototype, "addFollowups");

    const response = {
      markdown: sandbox.stub(),
      button: sandbox.stub(),
    };
    const token = new CancellationToken();
    await officeNextStepCommandHandler(
      {} as vscode.ChatRequest,
      {} as vscode.ChatContext,
      response as unknown as vscode.ChatResponseStream,
      token
    );
    chai.assert.isTrue(response.markdown.calledThrice);
    chai.assert.isTrue(response.button.calledThrice);
    chai.assert.isTrue(followupProviderStub.calledOnce);
  });
});

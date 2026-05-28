import * as sinon from "sinon";
import * as chai from "chai";
import * as vscode from "vscode";
import * as localizeUtils from "../../src/utils/localizeUtils";
import fs from "fs-extra";
import * as globalVariables from "../../src/globalVariables";
import * as projectChecker from "../../src/utils/projectChecker";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { SystemError, UserError } from "@microsoft/teamsfx-api";
import { showError } from "../../src/error/common";
import { TelemetryEvent } from "../../src/telemetry/extTelemetryEvents";
import { RecommendedOperations } from "../../src/debug/common/debugConstants";

describe("common", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("showError", async () => {
    sandbox.stub(localizeUtils, "localize").returns("");
    const showErrorMessageStub = sandbox
      .stub(vscode.window, "showErrorMessage")
      .callsFake((title: string, button: any) => {
        return Promise.resolve(button);
      });
    const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(vscode.commands, "executeCommand");
    const error = new UserError("test source", "test name", "test message", "test displayMessage");
    error.helpLink = "test helpLink";

    await showError(error);

    chai.assert.isTrue(
      sendTelemetryEventStub.calledWith(TelemetryEvent.ClickGetHelp, {
        "error-code": "test source.test name",
        "err-message": "test displayMessage",
        "help-link": "test helpLink",
      })
    );
  });

  it("showError with test tool button click", async () => {
    sandbox.stub(localizeUtils, "localize").returns("");
    const showErrorMessageStub = sandbox
      .stub(vscode.window, "showErrorMessage")
      .callsFake((title: string, button: any) => {
        return Promise.resolve(button);
      });
    const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(vscode.commands, "executeCommand");
    const error = new UserError("test source", "test name", "test message", "test displayMessage");
    error.recommendedOperation = "debug-in-test-tool";
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("path"));
    sandbox.stub(fs, "pathExistsSync").returns(true);

    await showError(error);

    chai.assert.isFalse(
      sendTelemetryEventStub.calledWith(TelemetryEvent.ClickGetHelp, {
        "error-code": "test source.test name",
        "err-message": "test displayMessage",
        "help-link": "test helpLink",
      })
    );
  });

  it("showError - similar issues", async () => {
    sandbox
      .stub(vscode.window, "showErrorMessage")
      .callsFake((title: string, button: unknown, ...items: vscode.MessageItem[]) => {
        return Promise.resolve(items[0]);
      });
    const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");
    const error = new SystemError("Core", "DecryptionError", "test");

    await showError(error);

    chai.assert.isTrue(sendTelemetryEventStub.called);
    chai.assert.isTrue(executeCommandStub.called);
  });

  [
    {
      type: "user error",
      buildError: () => {
        const error = new UserError(
          "test source",
          "test name",
          "test message",
          "test displayMessage"
        );
        error.helpLink = "test helpLink";
        error.recommendedOperation = RecommendedOperations.DebugInTestTool;

        return error;
      },
      buttonNum: 2,
    },
    {
      type: "system error",
      buildError: () => {
        const error = new SystemError(
          "test source",
          "test name",
          "test message",
          "test displayMessage"
        );
        error.recommendedOperation = RecommendedOperations.DebugInTestTool;
        return error;
      },
      buttonNum: 3,
    },
  ].forEach(({ type, buildError, buttonNum }) => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it(`showError - ${type} - recommend test tool`, async () => {
      sandbox.stub(localizeUtils, "localize").returns("");
      const showErrorMessageStub = sandbox.stub(vscode.window, "showErrorMessage");
      sandbox.stub(projectChecker, "isTestToolEnabledProject").returns(true);
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("path"));
      sandbox.stub(vscode.commands, "executeCommand");
      const error = buildError();
      await showError(error);
      chai.assert.equal(showErrorMessageStub.firstCall.args.length, buttonNum + 1);
    });
  });
});

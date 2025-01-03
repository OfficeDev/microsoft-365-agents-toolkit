import * as sinon from "sinon";
import * as chai from "chai";
import * as vscode from "vscode";
import * as localizeUtils from "../../src/utils/localizeUtils";
import fs from "fs-extra";
import * as globalVariables from "../../src/globalVariables";
import * as projectChecker from "../../src/utils/projectChecker";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { SystemError, UserError } from "@microsoft/teamsfx-api";
import { notifyOutputTroubleshoot, showError } from "../../src/error/common";
import { TelemetryEvent } from "../../src/telemetry/extTelemetryEvents";
import { RecommendedOperations } from "../../src/debug/common/debugConstants";
import { featureFlagManager } from "@microsoft/teamsfx-core";
import { MaximumNotificationOutputTroubleshootCount } from "../../src/constants";

describe("common", async () => {
  const sandbox = sinon.createSandbox();
  let clock: sinon.SinonFakeTimers;

  afterEach(() => {
    sandbox.restore();
    if (clock) {
      clock.restore();
    }
  });

  it("showError", async () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
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
    await showErrorMessageStub.firstCall.returnValue;

    chai.assert.isTrue(
      sendTelemetryEventStub.calledWith(TelemetryEvent.ClickGetHelp, {
        "error-code": "test source.test name",
        "err-message": "test displayMessage",
        "help-link": "test helpLink",
      })
    );
  });

  it("showError - if user does not click any button", async () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
    sandbox.stub(localizeUtils, "localize").returns("");
    const showErrorMessageStub = sandbox
      .stub(vscode.window, "showErrorMessage")
      .callsFake((title: string, button: any) => {
        return Promise.resolve(undefined);
      });
    const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(vscode.commands, "executeCommand");
    const error = new UserError("test source", "test name", "test message", "test displayMessage");
    error.helpLink = "test helpLink";

    await showError(error);
    await showErrorMessageStub.firstCall.returnValue;

    chai.assert.isFalse(
      sendTelemetryEventStub.calledWith(TelemetryEvent.ClickGetHelp, {
        "error-code": "test source.test name",
        "err-message": "test displayMessage",
        "help-link": "test helpLink",
      })
    );
  });

  it("showError with test tool button click", async () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
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
    await showErrorMessageStub.firstCall.returnValue;

    chai.assert.isFalse(
      sendTelemetryEventStub.calledWith(TelemetryEvent.ClickGetHelp, {
        "error-code": "test source.test name",
        "err-message": "test displayMessage",
        "help-link": "test helpLink",
      })
    );
  });

  it("showError - similar issues", async () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
    const showErrorMessageStub = sandbox
      .stub(vscode.window, "showErrorMessage")
      .callsFake((title: string, button: unknown, ...items: vscode.MessageItem[]) => {
        return Promise.resolve(items[0]);
      });

    const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");
    const error = new SystemError("Core", "DecryptionError", "test");

    await showError(error);
    await showErrorMessageStub.firstCall.returnValue;

    chai.assert.isTrue(executeCommandStub.called);
  });

  it("showError - similar issues and no button clicked", async () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
    const showErrorMessageStub = sandbox
      .stub(vscode.window, "showErrorMessage")
      .callsFake((title: string, button: unknown, ...items: vscode.MessageItem[]) => {
        return Promise.resolve(undefined);
      });
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");
    const error = new SystemError("Core", "DecryptionError", "test");

    await showError(error);
    await showErrorMessageStub.firstCall.returnValue;

    chai.assert.isTrue(executeCommandStub.notCalled);
  });

  describe("notify user to troubleshoot output with Teams Agent", async () => {
    let showInformationMessageStub: sinon.SinonStub;
    let showErrorMessageStub: sinon.SinonStub;
    beforeEach(() => {
      showInformationMessageStub = sandbox.stub(vscode.window, "showInformationMessage");
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      showErrorMessageStub = sandbox
        .stub(vscode.window, "showErrorMessage")
        .callsFake((title: string, button: any) => {
          return Promise.resolve(button);
        });
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      clock = sandbox.useFakeTimers();
    });

    afterEach(() => {
      globalVariables.setOutputTroubleshootNotificationCount(0);
      if (clock) {
        clock.restore();
      }
    });
    it("showError - notify user to troubleshoot output with Teams Agent", async () => {
      showInformationMessageStub.resolves("Open output panel");
      globalVariables.setOutputTroubleshootNotificationCount(0);
      sandbox.stub(vscode.commands, "executeCommand");
      const error = new UserError(
        "test source",
        "test name",
        "test message",
        "test displayMessage"
      );
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("path"));
      sandbox.stub(fs, "pathExistsSync").returns(true);

      const job = showError(error);
      await clock.tickAsync(4000);
      await job;
      await showErrorMessageStub.firstCall.returnValue;

      chai.assert.equal(globalVariables.outputTroubleshootNotificationCount, 1);
    });

    it("showError - not notify user to troubleshoot output with Teams Agent if reaches limit", async () => {
      globalVariables.setOutputTroubleshootNotificationCount(3);
      sandbox.stub(vscode.commands, "executeCommand");
      const error = new UserError(
        "test source",
        "test name",
        "test message",
        "test displayMessage"
      );
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("path"));
      sandbox.stub(fs, "pathExistsSync").returns(true);

      await showError(error);
      await showErrorMessageStub.firstCall.returnValue;

      chai.assert.equal(globalVariables.outputTroubleshootNotificationCount, 3);
      chai.assert.isTrue(showErrorMessageStub.calledOnce);
    });

    it("showError - not notify user to troubleshoot output with Teams Agent if userCancelError", async () => {
      globalVariables.setOutputTroubleshootNotificationCount(0);

      sandbox.stub(vscode.commands, "executeCommand");
      const error = new UserError(
        "test source",
        "User Cancel",
        "test message",
        "test displayMessage"
      );
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("path"));
      sandbox.stub(fs, "pathExistsSync").returns(true);

      await showError(error);

      chai.assert.equal(globalVariables.outputTroubleshootNotificationCount, 0);
      chai.assert.isFalse(showErrorMessageStub.called);
    });

    it("should execute command when user selects 'Open output panel'", async () => {
      showInformationMessageStub.callsFake((title: string, button: any) => {
        return Promise.resolve(button);
      });
      const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();

      const job = notifyOutputTroubleshoot("testErrorCode");
      await clock.tickAsync(4000);
      await job;
      await showInformationMessageStub.firstCall.returnValue;

      chai.assert.isTrue(executeCommandStub.calledOnceWith("fx-extension.showOutputChannel"));
    });
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
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      sandbox.stub(localizeUtils, "localize").returns("");
      const showErrorMessageStub = sandbox
        .stub(vscode.window, "showErrorMessage")
        .callsFake((title: string, button: any) => {
          return Promise.resolve(button);
        });
      sandbox.stub(projectChecker, "isTestToolEnabledProject").returns(true);
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("path"));
      sandbox.stub(vscode.commands, "executeCommand");
      const error = buildError();
      await showError(error);
      await showErrorMessageStub.firstCall.returnValue;
      chai.assert.equal(showErrorMessageStub.firstCall.args.length, buttonNum + 1);
    });

    it(`showError - ${type} - recommend troubleshoot`, async () => {
      clock = sandbox.useFakeTimers();
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      globalVariables.setOutputTroubleshootNotificationCount(
        MaximumNotificationOutputTroubleshootCount
      );
      const showErrorMessageStub = sandbox
        .stub(vscode.window, "showErrorMessage")
        .callsFake((title: string, button: any) => {
          return Promise.resolve(button);
        });
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      sandbox.stub(localizeUtils, "localize").returns("");
      sandbox.stub(projectChecker, "isTestToolEnabledProject").returns(true);
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("path"));
      sandbox.stub(vscode.commands, "executeCommand");
      const error = buildError();
      const job = showError(error);
      await clock.tickAsync(4000);
      await job;
      await showErrorMessageStub.firstCall.returnValue;

      if (type == "system error") {
        chai.assert.equal(showErrorMessageStub.firstCall.args.length, buttonNum + 1);
      } else {
        chai.assert.equal(showErrorMessageStub.firstCall.args.length, buttonNum + 2);
      }
    });

    it(`showError - ${type} - recommend troubleshoot only`, async () => {
      clock = sandbox.useFakeTimers();
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      globalVariables.setOutputTroubleshootNotificationCount(
        MaximumNotificationOutputTroubleshootCount
      );
      const showErrorMessageStub = sandbox
        .stub(vscode.window, "showErrorMessage")
        .callsFake((title: string, button: any) => {
          return Promise.resolve(button);
        });
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      sandbox.stub(localizeUtils, "localize").returns("");
      sandbox.stub(projectChecker, "isTestToolEnabledProject").returns(false);
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("path"));
      sandbox.stub(vscode.commands, "executeCommand");
      const error = buildError();
      const job = showError(error);
      await clock.tickAsync(4000);
      await job;
      await showErrorMessageStub.firstCall.returnValue;

      chai.assert.equal(showErrorMessageStub.firstCall.args.length, buttonNum + 1);
    });
  });
});

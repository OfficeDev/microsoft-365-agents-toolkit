import * as sinon from "sinon";
import * as chai from "chai";
import * as vscode from "vscode";
import * as localizeUtils from "../../src/utils/localizeUtils";
import fs from "fs-extra";
import * as globalVariables from "../../src/globalVariables";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { SystemError, UserError } from "@microsoft/teamsfx-api";
import {
  isLoginFailureError,
  notifyOutputTroubleshoot,
  showError,
  wrapError,
} from "../../src/error/common";
import { TelemetryEvent } from "../../src/telemetry/extTelemetryEvents";
import { RecommendedOperations } from "../../src/debug/common/debugConstants";
import { featureFlagManager, GraphClient, FeatureFlagName } from "@microsoft/teamsfx-core";
import { MaximumNotificationOutputTroubleshootCount } from "../../src/constants";
import * as tools from "@microsoft/teamsfx-core/build/common/tools";

describe("common", async () => {
  const sandbox = sinon.createSandbox();
  let clock: sinon.SinonFakeTimers;

  afterEach(() => {
    sandbox.restore();
    if (clock) {
      clock.restore();
    }
  });

  beforeEach(() => {
    sandbox.stub(GraphClient.prototype, "GetTeamsAppSettingsAsync").resolves({
      sandboxingConfiguration: {
        isSideloadingEnabled: true,
        sensitivityLabelUsedToIdentifySandboxedContainers: "0fcfd0ff-1cda-407e-bc2b-a350307bd1d5",
      },
    });
  });

  it("showError", async () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
    sandbox.stub(localizeUtils, "localize").returns("");
    const showErrorMessageStub = sandbox
      .stub(vscode.window, "showErrorMessage")
      .callsFake((title: string, button: any) => {
        return Promise.resolve(button);
      });
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(vscode.commands, "executeCommand");
    const error = new UserError("test source", "test name", "test message", "test displayMessage");
    error.helpLink = "test helpLink";

    await showError(error);
    await showErrorMessageStub.firstCall.returnValue;

    // "Get Help" button has been removed; only the troubleshoot button is shown
    // when no other recommendations apply, so the call should still happen with
    // exactly one button.
    chai.assert.isTrue(showErrorMessageStub.calledOnce);
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

      // The auto-notify-after-error info message has been removed; the count
      // should remain unchanged.
      chai.assert.equal(globalVariables.outputTroubleshootNotificationCount, 0);
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
      // "Get Help" button removed → 1 button (runTestTool)
      buttonNum: 1,
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
      // System error path: runTestTool + issue + similarIssues = 3 buttons
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
      sandbox.stub(tools, "isTestToolEnabledProject").returns(true);
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
      sandbox.stub(featureFlagManager, "getBooleanValue").callsFake((featureFlag: any) => {
        if (featureFlag.name == FeatureFlagName.SandBoxedTeam) {
          return false;
        } else {
          return true;
        }
      });
      sandbox.stub(localizeUtils, "localize").returns("");
      sandbox.stub(tools, "isTestToolEnabledProject").returns(true);
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
      sandbox.stub(tools, "isTestToolEnabledProject").returns(true);
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
        // User-error helpLink branch with sandbox + troubleshoot:
        // [troubleshoot, runSandbox] = 2 buttons + title = 3 args
        chai.assert.equal(showErrorMessageStub.firstCall.args.length, buttonNum + 2);
      }
    });
  });

  describe("button click handlers", () => {
    it("runTestTool button: opens Microsoft 365 Agents Playground debug picker", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      sandbox.stub(localizeUtils, "localize").returns("");
      sandbox.stub(tools, "isTestToolEnabledProject").returns(true);
      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("path"));
      const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();

      const error = new UserError(
        "test source",
        "test name",
        "test message",
        "test displayMessage"
      );
      error.recommendedOperation = RecommendedOperations.DebugInTestTool;

      sandbox
        .stub(vscode.window, "showErrorMessage")
        .callsFake(async (title: string, ...buttons: any[]) => {
          // The first (and only) button in the user-error fallback branch is runTestTool.
          const button = buttons[0];
          await button.run();
          return button;
        });

      await showError(error);

      chai.assert.isTrue(
        executeCommandStub.calledWith(
          "workbench.action.quickOpen",
          "debug Debug in Microsoft 365 Agents Playground"
        )
      );
      chai.assert.isTrue(sendTelemetryEventStub.calledWith(TelemetryEvent.MessageDebugInTestTool));
    });

    it("runSandbox button: opens sandbox debug picker", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").callsFake((flag: any) => {
        // SandBoxedTeam => true (so recommendSandbox=true), all others => false
        return flag.name === FeatureFlagName.SandBoxedTeam;
      });
      sandbox.stub(localizeUtils, "localize").returns("");
      const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();

      const error = new UserError(
        "test source",
        "test name",
        "test message",
        "test displayMessage"
      );
      error.helpLink = "test helpLink";

      sandbox
        .stub(vscode.window, "showErrorMessage")
        .callsFake(async (title: string, ...buttons: any[]) => {
          const button = buttons[0];
          await button.run();
          return button;
        });

      await showError(error);

      chai.assert.isTrue(
        executeCommandStub.calledWith(
          "workbench.action.quickOpen",
          "debug Debug in sandbox in Teams (Edge)"
        )
      );
      chai.assert.isTrue(sendTelemetryEventStub.calledWith(TelemetryEvent.MessageDebugInSandbox));
    });

    it("issue button: opens GitHub bug report URL", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      sandbox.stub(localizeUtils, "localize").returns("");
      const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();

      const error = new SystemError("Core", "TestSystemError", "test message");

      sandbox
        .stub(vscode.window, "showErrorMessage")
        .callsFake(async (title: string, ...buttons: any[]) => {
          // System-error path with shouldRecommendTeamsAgent=false and no recommendations
          // => buttons = [issue, similarIssues]. Click the first one.
          const button = buttons[0];
          await button.run();
          return button;
        });

      await showError(error);

      chai.assert.isTrue(executeCommandStub.calledWith("vscode.open", sinon.match.any));
    });
  });

  describe("wrapError", () => {
    it("returns the error directly when input is a UserError", () => {
      const original = new UserError("src", "name", "message");
      const result = wrapError(original);
      chai.assert.isTrue(result.isErr());
      if (result.isErr()) {
        chai.assert.strictEqual(result.error, original);
      }
    });

    it("returns the error directly when input is a SystemError", () => {
      const original = new SystemError("src", "name", "message");
      const result = wrapError(original);
      chai.assert.isTrue(result.isErr());
      if (result.isErr()) {
        chai.assert.strictEqual(result.error, original);
      }
    });

    it("wraps a plain Error in a SystemError", () => {
      const original = new Error("boom");
      const result = wrapError(original);
      chai.assert.isTrue(result.isErr());
      if (result.isErr()) {
        chai.assert.instanceOf(result.error, SystemError);
      }
    });
  });

  describe("isLoginFailureError", () => {
    it("returns true when the message contains the login failure marker", () => {
      const e = new UserError("src", "name", "Cannot get user login information from cache");
      chai.assert.isTrue(isLoginFailureError(e));
    });

    it("returns false for unrelated errors", () => {
      const e = new UserError("src", "name", "Something else went wrong");
      chai.assert.isFalse(isLoginFailureError(e));
    });

    it("returns false when message is empty", () => {
      const e = new UserError("src", "name", "");
      chai.assert.isFalse(isLoginFailureError(e));
    });
  });
});

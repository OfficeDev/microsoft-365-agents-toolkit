import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";

import VsCodeLogInstance from "../../src/commonlib/log";
import * as handlers from "../../src/handlers/copilotChatHandlers";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import * as extTelemetryEvents from "../../src/telemetry/extTelemetryEvents";
import * as versionUtils from "../../src/utils/versionUtil";
import * as globalState from "@microsoft/teamsfx-core/build/common/globalState";
import { localize } from "../../src/utils/localizeUtils";
import * as vsc_ui from "../../src/qm/vsc_ui";
import { err, ok, SystemError } from "@microsoft/teamsfx-api";

after(() => {
  sinon.restore();
});

describe("invokeTeamsAgent", async () => {
  const sandbox = sinon.createSandbox();
  let clock: sinon.SinonFakeTimers;

  afterEach(() => {
    sandbox.restore();
    if (clock) {
      clock.restore();
    }
  });

  beforeEach(() => {
    sandbox.stub(ExtTelemetry, "dispose");
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
    sandbox.stub(VsCodeLogInstance, "outputChannel").value({
      name: "name",
      append: (value: string) => {},
      appendLine: (value: string) => {},
      replace: (value: string) => {},
      clear: () => {},
      show: (...params: any[]) => {},
      hide: () => {},
      dispose: () => {},
    });
    sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
  });

  it("no need to install Github Copilot", async () => {
    sandbox.stub(globalState, "globalStateGet").resolves(true);
    sandbox.stub(vscode.extensions, "getExtension").returns({ name: "github.copilot" } as any);
    sandbox.stub(vscode.commands, "executeCommand").resolves();

    const res = await handlers.invokeTeamsAgent([
      extTelemetryEvents.TelemetryTriggerFrom.CreateAppQuestionFlow,
    ]);

    chai.assert.isTrue(res.isOk());
  });

  it("install Github Copilot and invoke Teams Agent", async () => {
    sandbox.stub(globalState, "globalStateGet").resolves(true);
    clock = sandbox.useFakeTimers();
    sandbox.stub(versionUtils, "isVSCodeInsiderVersion").returns(true);
    sandbox
      .stub(vscode.extensions, "getExtension")
      .onFirstCall()
      .returns(undefined)
      .onSecondCall()
      .returns({ name: "github.copilot" } as any);
    const commandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();
    sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves(
        localize("teamstoolkit.handlers.askInstallCopilot.install") as unknown as vscode.MessageItem
      );

    const job = handlers.invokeTeamsAgent([extTelemetryEvents.TelemetryTriggerFrom.TreeView]);
    await clock.tickAsync(6000);
    const res = await job;

    if (res.isErr()) {
      console.log(res.error);
    }

    chai.assert.isTrue(res.isOk());
    chai.assert.equal(commandStub.callCount, 3);
    chai.assert.isTrue(
      (commandStub.getCall(2).args[1].query as string).startsWith("@teamsapp Use ")
    );
  });

  it("install Github Copilot and invoke Teams Agent - WalkThrough - 1", async () => {
    sandbox.stub(globalState, "globalStateGet").resolves(true);
    clock = sandbox.useFakeTimers();
    sandbox.stub(versionUtils, "isVSCodeInsiderVersion").returns(true);
    sandbox
      .stub(vscode.extensions, "getExtension")
      .onFirstCall()
      .returns(undefined)
      .onSecondCall()
      .returns({ name: "github.copilot" } as any);
    const commandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();
    sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves(
        localize("teamstoolkit.handlers.askInstallCopilot.install") as unknown as vscode.MessageItem
      );

    const job = handlers.invokeTeamsAgent([
      extTelemetryEvents.TelemetryTriggerFrom.WalkThroughIntroduction,
    ]);
    await clock.tickAsync(6000);
    const res = await job;
    chai.assert.isTrue(res.isOk());
    chai.assert.equal(commandStub.callCount, 3);
    chai.assert.isTrue(
      (commandStub.getCall(2).args[1].query as string).startsWith(
        "@teamsapp What is notification bot in Teams"
      )
    );
  });

  it("install Github Copilot and invoke Teams Agent - WalkThrough - 2", async () => {
    sandbox.stub(globalState, "globalStateGet").resolves(true);
    clock = sandbox.useFakeTimers();
    sandbox.stub(versionUtils, "isVSCodeInsiderVersion").returns(true);
    sandbox
      .stub(vscode.extensions, "getExtension")
      .onFirstCall()
      .returns(undefined)
      .onSecondCall()
      .returns({ name: "github.copilot" } as any);
    const commandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();
    sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves(
        localize("teamstoolkit.handlers.askInstallCopilot.install") as unknown as vscode.MessageItem
      );

    const job = handlers.invokeTeamsAgent([
      extTelemetryEvents.TelemetryTriggerFrom.WalkThroughCreate,
    ]);
    await clock.tickAsync(6000);
    const res = await job;
    chai.assert.isTrue(res.isOk());
    chai.assert.equal(commandStub.callCount, 3);
    chai.assert.isTrue(
      (commandStub.getCall(2).args[1].query as string).startsWith(
        "@teamsapp How to create notification bot with Teams Toolkit?"
      )
    );
  });

  it("install Github Copilot and invoke Teams Agent - WalkThrough - 3", async () => {
    sandbox.stub(globalState, "globalStateGet").resolves(true);
    clock = sandbox.useFakeTimers();
    sandbox.stub(versionUtils, "isVSCodeInsiderVersion").returns(true);
    sandbox
      .stub(vscode.extensions, "getExtension")
      .onFirstCall()
      .returns(undefined)
      .onSecondCall()
      .returns({ name: "github.copilot" } as any);
    const commandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();
    sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves(
        localize("teamstoolkit.handlers.askInstallCopilot.install") as unknown as vscode.MessageItem
      );

    const job = handlers.invokeTeamsAgent([
      extTelemetryEvents.TelemetryTriggerFrom.WalkThroughWhatIsNext,
    ]);
    await clock.tickAsync(6000);
    const res = await job;
    chai.assert.isTrue(res.isOk());
    chai.assert.equal(commandStub.callCount, 3);
    chai.assert.isTrue(
      (commandStub.getCall(2).args[1].query as string).startsWith(
        "@teamsapp How do I customize and extend the notification bot app template created by Teams Toolkit?"
      )
    );
  });

  it("install Github Copilot and invoke Teams Agent - WalkThrough - 4", async () => {
    sandbox.stub(globalState, "globalStateGet").resolves(true);
    clock = sandbox.useFakeTimers();
    sandbox.stub(versionUtils, "isVSCodeInsiderVersion").returns(true);
    sandbox
      .stub(vscode.extensions, "getExtension")
      .onFirstCall()
      .returns(undefined)
      .onSecondCall()
      .returns({ name: "github.copilot" } as any);
    const commandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();
    sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves(
        localize("teamstoolkit.handlers.askInstallCopilot.install") as unknown as vscode.MessageItem
      );

    const job = handlers.invokeTeamsAgent([
      extTelemetryEvents.TelemetryTriggerFrom.WalkThroughIntelligentAppsIntroduction,
    ]);
    await clock.tickAsync(6000);
    const res = await job;
    chai.assert.isTrue(res.isOk());
    chai.assert.equal(commandStub.callCount, 3);
    chai.assert.isTrue(
      (commandStub.getCall(2).args[1].query as string).startsWith(
        "@teamsapp What is declarative agent for Microsoft 365 Copilot?"
      )
    );
  });

  it("install Github Copilot and invoke Teams Agent - WalkThrough - 5", async () => {
    sandbox.stub(globalState, "globalStateGet").resolves(true);
    clock = sandbox.useFakeTimers();
    sandbox.stub(versionUtils, "isVSCodeInsiderVersion").returns(true);
    sandbox
      .stub(vscode.extensions, "getExtension")
      .onFirstCall()
      .returns(undefined)
      .onSecondCall()
      .returns({ name: "github.copilot" } as any);
    const commandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();
    sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves(
        localize("teamstoolkit.handlers.askInstallCopilot.install") as unknown as vscode.MessageItem
      );

    const job = handlers.invokeTeamsAgent([
      extTelemetryEvents.TelemetryTriggerFrom.WalkThroughIntelligentAppsCreate,
    ]);
    await clock.tickAsync(6000);
    const res = await job;
    chai.assert.isTrue(res.isOk());
    chai.assert.equal(commandStub.callCount, 3);
    chai.assert.isTrue(
      (commandStub.getCall(2).args[1].query as string).startsWith(
        "@teamsapp How to create declarative agent with Teams Toolkit?"
      )
    );
  });

  it("install Github Copilot, wait and invoke Teams Agent", async () => {
    sandbox.stub(globalState, "globalStateGet").resolves(true);
    clock = sandbox.useFakeTimers();
    sandbox.stub(versionUtils, "isVSCodeInsiderVersion").returns(true);
    sandbox
      .stub(vscode.extensions, "getExtension")
      .onFirstCall()
      .returns(undefined)
      .onSecondCall()
      .returns(undefined)
      .onThirdCall()
      .returns({ name: "github.copilot" } as any);
    const commandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();
    sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves(
        localize("teamstoolkit.handlers.askInstallCopilot.install") as unknown as vscode.MessageItem
      );

    const job = handlers.invokeTeamsAgent();
    await clock.tickAsync(6000);
    const res = await job;

    chai.assert.isTrue(res.isOk());
    chai.assert.equal(commandStub.callCount, 3);
  });

  it("Install github copilot extension error", async () => {
    sandbox.stub(globalState, "globalStateGet").resolves(true);
    sandbox.stub(versionUtils, "isVSCodeInsiderVersion").returns(true);
    sandbox.stub(vscode.extensions, "getExtension").onFirstCall().returns(undefined);
    const commandStub = sandbox
      .stub(vscode.commands, "executeCommand")
      .callsFake(async (command: string) => {
        if (command === "workbench.extensions.installExtension") {
          throw new Error("Install Error");
        } else {
          return {};
        }
      });
    sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves(
        localize("teamstoolkit.handlers.askInstallCopilot.install") as unknown as vscode.MessageItem
      );
    sandbox.stub(VsCodeLogInstance, "error").resolves();

    const res = await handlers.invokeTeamsAgent();

    chai.assert.isTrue(res.isErr());
    if (res.isErr()) {
      chai.assert.equal(res.error.source, "install-copilot-chat");
    }
    chai.assert.equal(commandStub.callCount, 1);
  });

  it("Install github copilot extension cancel", async () => {
    sandbox.stub(globalState, "globalStateGet").resolves(true);
    sandbox.stub(versionUtils, "isVSCodeInsiderVersion").returns(true);
    const loggerStub = sandbox.stub(VsCodeLogInstance, "error").resolves();
    sandbox
      .stub(vscode.extensions, "getExtension")
      .onFirstCall()
      .returns(undefined)
      .onSecondCall()
      .returns(undefined)
      .onThirdCall()
      .returns({ name: "github.copilot" } as any);
    const commandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();

    sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves("Cancel" as unknown as vscode.MessageItem);

    const res = await handlers.invokeTeamsAgent();

    chai.assert.isTrue(res.isErr());
    if (res.isErr()) {
      chai.assert.equal(res.error.name, "UserCancelError");
    }
    chai.assert.equal(commandStub.callCount, 0);
    chai.assert.equal(loggerStub.callCount, 0);
  });

  it("Verify installation error", async () => {
    sandbox.stub(globalState, "globalStateGet").resolves(true);
    clock = sandbox.useFakeTimers();
    sandbox.stub(versionUtils, "isVSCodeInsiderVersion").returns(true);
    sandbox.stub(vscode.extensions, "getExtension").returns(undefined);
    const commandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();
    sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves(
        localize("teamstoolkit.handlers.askInstallCopilot.install") as unknown as vscode.MessageItem
      );

    const job = handlers.invokeTeamsAgent();
    await clock.tickAsync(30000);
    const res = await job;

    chai.assert.isTrue(res.isErr());
    if (res.isErr()) {
      chai.assert.equal(res.error.name, "CannotVerifyGithubCopilotChat");
    }
    chai.assert.equal(commandStub.callCount, 1);
  });

  it("invoke Copilot chat error", async () => {
    sandbox.stub(globalState, "globalStateGet").resolves(true);
    sandbox.stub(versionUtils, "isVSCodeInsiderVersion").returns(true);
    sandbox.stub(vscode.extensions, "getExtension").returns({ name: "github.copilot" } as any);
    const commandStub = sandbox
      .stub(vscode.commands, "executeCommand")
      .callsFake(async (command: string) => {
        if (command === "workbench.panel.chat.view.copilot.focus") {
          throw new Error("Install Error");
        } else {
          return {};
        }
      });
    sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves(
        localize("teamstoolkit.handlers.askInstallCopilot.install") as unknown as vscode.MessageItem
      );
    const loggerError = sandbox.stub(VsCodeLogInstance, "error").resolves();

    const res = await handlers.invokeTeamsAgent();

    chai.assert.isTrue(res.isErr());
    if (res.isErr()) {
      chai.assert.equal(res.error.source, "openCopilotChat");
    }
    chai.assert.equal(commandStub.callCount, 1);
    chai.assert.equal(loggerError.callCount, 2);
  });

  it("need to show notification of installing @teamsapp", async () => {
    sandbox.stub(globalState, "globalStateGet").resolves(false);
    sandbox.stub(vscode.extensions, "getExtension").returns({ name: "github.copilot" } as any);
    sandbox.stub(vscode.commands, "executeCommand").resolves();

    sandbox
      .stub(vscode.window, "showInformationMessage")
      .returns(
        Promise.resolve(
          localize(
            "teamstoolkit.handlers.askInstallTeamsAgent.install"
          ) as unknown as vscode.MessageItem
        )
      );
    sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
    const res = await handlers.invokeTeamsAgent([
      extTelemetryEvents.TelemetryTriggerFrom.CreateAppQuestionFlow,
    ]);
    chai.assert.isTrue(res.isOk());
  });

  describe("handleInstallTeamsAgentSelection", async () => {
    it("open url", async () => {
      const openUrlStub = sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      await handlers.handleInstallTeamsAgentSelection(
        localize("teamstoolkit.handlers.askInstallTeamsAgent.install"),
        { key: "value" }
      );

      chai.assert.isTrue(openUrlStub.called);
    });

    it("confirm install", async () => {
      const stub = sandbox.stub(globalState, "globalStateUpdate").resolves(undefined);
      await handlers.handleInstallTeamsAgentSelection(
        localize("teamstoolkit.handlers.askInstallTeamsAgent.confirmInstall"),
        { key: "value" }
      );

      chai.assert.isTrue(stub.called);
    });

    it("open url error", async () => {
      const openUrlStub = sandbox
        .stub(vsc_ui.VS_CODE_UI, "openUrl")
        .resolves(err(new SystemError("openUrl", "openUrlError", "", "")));
      const logError = sandbox.stub(VsCodeLogInstance, "error").resolves();
      await handlers.handleInstallTeamsAgentSelection(
        localize("teamstoolkit.handlers.askInstallTeamsAgent.install"),
        { key: "value" }
      );

      chai.assert.isTrue(openUrlStub.called);
      chai.assert.isTrue(logError.called);
    });

    it("cancel", async () => {
      const openUrlStub = sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      await handlers.handleInstallTeamsAgentSelection(undefined, { key: "value" });

      chai.assert.isTrue(openUrlStub.notCalled);
    });
  });

  describe("troubleshootSelectedText", async () => {
    it("can invoke teams agent", async () => {
      sandbox.stub(vscode.window, "activeTextEditor").value({
        selection: "current select",
        document: {
          getText: (selection: vscode.Selection) => "current select",
        },
      } as any);
      sandbox.stub(globalState, "globalStateGet").resolves(true);
      sandbox.stub(vscode.extensions, "getExtension").returns({ name: "github.copilot" } as any);
      sandbox.stub(vscode.commands, "executeCommand").resolves();
      const res = await handlers.troubleshootSelectedText();
      if (res.isErr()) {
        console.log(res.error);
      }
      chai.assert.isTrue(res.isOk());
    });

    it("no active text", async () => {
      sandbox.stub(vscode.window, "activeTextEditor").value(undefined);
      const res = await handlers.troubleshootSelectedText();
      chai.assert.isTrue(res.isErr());
    });

    it("error", async () => {
      sandbox.stub(vscode.window, "activeTextEditor").value({
        selection: "current select",
        document: {
          getText: (selection: vscode.Selection) => "current select",
        },
      } as any);
      sandbox.stub(globalState, "globalStateGet").resolves(true);
      sandbox.stub(versionUtils, "isVSCodeInsiderVersion").returns(true);
      sandbox.stub(vscode.extensions, "getExtension").onFirstCall().returns(undefined);
      const res = await handlers.troubleshootSelectedText();
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        console.log(res.error);
      }
    });
  });

  describe("troubleshootError", async () => {
    it("can invoke teams agent", async () => {
      sandbox.stub(globalState, "globalStateGet").resolves(true);
      sandbox.stub(vscode.extensions, "getExtension").returns({ name: "github.copilot" } as any);
      sandbox.stub(vscode.commands, "executeCommand").resolves();

      const currentError = new SystemError("test", "test", "test", "test");
      const res = await handlers.troubleshootError(["triggerFrom", currentError]);
      chai.assert.isTrue(res.isOk());
    });

    it("missing args", async () => {
      const res = await handlers.troubleshootError([]);
      const calledCommand = sandbox.stub(vscode.commands, "executeCommand").resolves();
      chai.assert.isTrue(res.isOk());
      chai.assert.isFalse(calledCommand.calledOnce);
    });

    it("error", async () => {
      sandbox.stub(globalState, "globalStateGet").resolves(true);
      sandbox.stub(versionUtils, "isVSCodeInsiderVersion").returns(true);
      sandbox.stub(vscode.extensions, "getExtension").onFirstCall().returns(undefined);
      sandbox.stub(vscode.commands, "executeCommand").callsFake(async (command: string) => {
        if (command === "workbench.extensions.installExtension") {
          throw new Error("Install Error");
        } else {
          return {};
        }
      });

      const currentError = new SystemError("test", "test", "test", "test");
      const res = await handlers.troubleshootError(["triggerFrom", currentError]);
      chai.assert.isTrue(res.isErr());
    });
  });
});

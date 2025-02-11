import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";

import VsCodeLogInstance from "../../src/commonlib/log";
import * as handlers from "../../src/handlers/copilotChatHandlers";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import * as extTelemetryEvents from "../../src/telemetry/extTelemetryEvents";
import * as versionUtils from "../../src/utils/versionUtil";
import * as globalState from "@microsoft/teamsfx-core/build/common/globalState";
import * as vsc_ui from "../../src/qm/vsc_ui";
import { err, ok, SystemError } from "@microsoft/teamsfx-api";
import { GlobalKey } from "../../src/constants";
import { TelemetryTriggerFrom } from "../../src/telemetry/extTelemetryEvents";

after(() => {
  sinon.restore();
});

describe("copilotChatHandler", async () => {
  const sandbox = sinon.createSandbox();
  let clock: sinon.SinonFakeTimers | undefined;
  let sendTelemetryErrorEventStub: sinon.SinonStub;

  afterEach(() => {
    sandbox.restore();
    if (clock) {
      clock.restore();
    }
  });

  beforeEach(() => {
    sandbox.stub(ExtTelemetry, "dispose");
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sendTelemetryErrorEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
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

  describe("openGithubCopilotChat", async () => {
    it("open without query success", async () => {
      const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();
      const res = await handlers.openGithubCopilotChat([
        extTelemetryEvents.TelemetryTriggerFrom.CreateAppQuestionFlow,
      ]);
      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(executeCommandStub.called);
    });

    it("open without query success", async () => {
      const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();
      const res = await handlers.openGithubCopilotChat([
        extTelemetryEvents.TelemetryTriggerFrom.CreateAppQuestionFlow,
        "test",
      ]);
      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(executeCommandStub.called);
    });

    it("open without query error", async () => {
      sandbox.stub(vscode.commands, "executeCommand").callsFake(async (command: string) => {
        if (command === "workbench.panel.chat.view.copilot.focus") {
          throw new Error("Install Error");
        } else {
          return {};
        }
      });

      sandbox.stub(VsCodeLogInstance, "error").resolves();

      const res = await handlers.openGithubCopilotChat();

      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.equal(res.error.source, "open-github-copilot-chat");
      }
    });

    it("open with query error", async () => {
      sandbox.stub(vscode.commands, "executeCommand").callsFake(async (command: string) => {
        if (command === "workbench.panel.chat.view.copilot.focus") {
          throw new Error("Install Error");
        } else {
          return {};
        }
      });

      sandbox.stub(VsCodeLogInstance, "error").resolves();

      const res = await handlers.openGithubCopilotChat([
        extTelemetryEvents.TelemetryTriggerFrom.CreateAppQuestionFlow,
        "test",
      ]);

      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.equal(res.error.source, "open-github-copilot-chat");
      }
    });
  });

  describe("installGithubCopilotChatExtension", async () => {
    it("no need to install Github Copilot", async () => {
      sandbox
        .stub(vscode.extensions, "getExtension")
        .returns({ name: "github.copilot-chat" } as any);
      sandbox.stub(vscode.commands, "executeCommand").resolves();

      const res = await handlers.installGithubCopilotChatExtension([
        extTelemetryEvents.TelemetryTriggerFrom.CreateAppQuestionFlow,
      ]);

      chai.assert.isTrue(res.isOk());
    });

    it("install Github Copilot successfully", async () => {
      sandbox.stub(versionUtils, "isVSCodeInsiderVersion").returns(true);
      const installStub = sandbox.stub(vscode.extensions, "getExtension").returns(undefined);

      const res = await handlers.installGithubCopilotChatExtension([
        extTelemetryEvents.TelemetryTriggerFrom.CreateAppQuestionFlow,
      ]);

      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(installStub.called);
    });

    it("Install github copilot extension error", async () => {
      sandbox.stub(versionUtils, "isVSCodeInsiderVersion").returns(true);
      sandbox.stub(vscode.extensions, "getExtension").returns(undefined);
      const commandStub = sandbox
        .stub(vscode.commands, "executeCommand")
        .callsFake(async (command: string) => {
          if (command === "workbench.extensions.installExtension") {
            throw new Error("Install Error");
          } else {
            return {};
          }
        });

      sandbox.stub(VsCodeLogInstance, "error").resolves();

      const res = await handlers.installGithubCopilotChatExtension();

      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.equal(res.error.source, "install-copilot-chat");
      }
      chai.assert.equal(commandStub.callCount, 1);
    });
  });

  describe("openInstallTeamsAgent", () => {
    it("should open URL successfully", async () => {
      const openUrlStub = sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      await handlers.openInstallTeamsAgent();
      chai.assert.isTrue(openUrlStub.calledOnce);
    });

    it("should handle URL opening failure", async () => {
      const error = new SystemError("test", "test", "test", "test");
      const openUrlStub = sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(err(error));
      const logErrorStub = sandbox.stub(VsCodeLogInstance, "error").resolves();
      await handlers.openInstallTeamsAgent();
      chai.assert.isTrue(openUrlStub.calledOnce);
      chai.assert.isTrue(logErrorStub.calledOnceWith(error.message));
    });
  });

  describe("markTeamsAgentInstallationDone", () => {
    it("should update global state successfully", async () => {
      const globalStateUpdateStub = sandbox.stub(globalState, "globalStateUpdate").resolves();
      await handlers.markTeamsAgentInstallationDone();
      chai.assert.isTrue(globalStateUpdateStub.calledOnceWith(GlobalKey.TeamsAgentInstalled, true));
    });

    it("should handle global state update failure", async () => {
      const error = new SystemError("test", "test", "test", "test");
      const globalStateUpdateStub = sandbox.stub(globalState, "globalStateUpdate").rejects(error);
      await handlers.markTeamsAgentInstallationDone();
      chai.assert.isTrue(globalStateUpdateStub.calledOnceWith(GlobalKey.TeamsAgentInstalled, true));
      chai.assert.isTrue(sendTelemetryErrorEventStub.calledOnce);
    });
  });

  describe("openTeamsAgentWalkthrough", () => {
    it("should execute command successfully", async () => {
      const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();
      await handlers.openTeamsAgentWalkthrough();
      chai.assert.isTrue(
        executeCommandStub.calledOnceWith("workbench.action.openWalkthrough", {
          category: "TeamsDevApp.ms-teams-vscode-extension#teamsAgentGetStarted",
        })
      );
    });
  });

  describe("invokeTeamsAgent", () => {
    it("open walkthrough successfully from treeview if GitHub copilot is not installed", async () => {
      const args = [TelemetryTriggerFrom.TreeView];
      sandbox.stub(globalState, "globalStateGet").resolves(true);
      sandbox.stub(vscode.extensions, "getExtension").returns(undefined);
      sandbox.stub(vscode.authentication, "getAccounts").resolves([
        {
          id: "someid",
          label: "",
        },
      ]);
      const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();

      const res = await handlers.invokeTeamsAgent(args);

      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        chai.assert.isFalse(res.value);
      }
      chai.assert.isTrue(executeCommandStub.calledOnce);
      chai.assert.isTrue(sendTelemetryErrorEventStub.notCalled);
    });

    it("open walkthrough successfully from treeview if not logged in", async () => {
      const args = [TelemetryTriggerFrom.TreeView];
      sandbox.stub(globalState, "globalStateGet").resolves(true);
      sandbox
        .stub(vscode.extensions, "getExtension")
        .returns({ name: "github.copilot-chat" } as any);
      sandbox.stub(vscode.authentication, "getAccounts").resolves([]);
      const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();

      const res = await handlers.invokeTeamsAgent(args);

      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        chai.assert.isFalse(res.value);
      }
      chai.assert.isTrue(executeCommandStub.calledOnce);
      chai.assert.isTrue(sendTelemetryErrorEventStub.notCalled);
    });

    it("open walkthrough successfully from treeview if @teamsapp not installed", async () => {
      const args = [TelemetryTriggerFrom.TreeView];
      sandbox.stub(globalState, "globalStateGet").resolves(false);
      sandbox
        .stub(vscode.extensions, "getExtension")
        .returns({ name: "github.copilot-chat" } as any);
      sandbox.stub(vscode.authentication, "getAccounts").resolves([
        {
          id: "someid",
          label: "",
        },
      ]);
      const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();

      const res = await handlers.invokeTeamsAgent(args);

      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        chai.assert.isFalse(res.value);
      }
      chai.assert.isTrue(executeCommandStub.calledOnce);
      chai.assert.isTrue(sendTelemetryErrorEventStub.notCalled);
    });

    it("invoke chat successfully from command palette", async () => {
      const args = [TelemetryTriggerFrom.CommandPalette];
      sandbox.stub(globalState, "globalStateGet").resolves(true);
      sandbox
        .stub(vscode.extensions, "getExtension")
        .returns({ name: "github.copilot-chat" } as any);
      sandbox.stub(vscode.authentication, "getAccounts").resolves([
        {
          id: "someid",
          label: "",
        },
      ]);
      const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();

      const res = await handlers.invokeTeamsAgent(args);

      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        chai.assert.isTrue(res.value);
      }
      chai.assert.isTrue(executeCommandStub.called);
      chai.assert.isTrue(sendTelemetryErrorEventStub.notCalled);
    });

    it("invoke chat successfully from unknown", async () => {
      const args = [TelemetryTriggerFrom.Unknow];
      sandbox.stub(globalState, "globalStateGet").resolves(true);
      sandbox
        .stub(vscode.extensions, "getExtension")
        .returns({ name: "github.copilot-chat" } as any);
      sandbox.stub(vscode.authentication, "getAccounts").resolves([
        {
          id: "someid",
          label: "",
        },
      ]);
      const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand").resolves();

      const res = await handlers.invokeTeamsAgent(args);

      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        chai.assert.isTrue(res.value);
      }
      chai.assert.isTrue(executeCommandStub.called);
      chai.assert.isTrue(sendTelemetryErrorEventStub.notCalled);
    });

    it("skip precheck and invoke chat error from WalkThrough", async () => {
      const args = [TelemetryTriggerFrom.WalkThrough];

      sandbox.stub(vscode.commands, "executeCommand").callsFake(async (command: string) => {
        if (command === "workbench.action.chat.open") {
          throw new Error("Error");
        } else {
          return {};
        }
      });

      const res = await handlers.invokeTeamsAgent(args);

      chai.assert.isTrue(res.isErr());
      if (res.isOk()) {
        chai.assert.isTrue(res.value);
      }
      chai.assert.isTrue(sendTelemetryErrorEventStub.called);
    });

    describe("invoke chat successfully from WalkThrough", async () => {
      const walkthroughTriggers = [
        TelemetryTriggerFrom.TeamsAgentWalkthroughCreate,
        TelemetryTriggerFrom.TeamsAgentWalkthroughExplore,
        TelemetryTriggerFrom.TeamsAgentWalkthroughTroubleshoot,
      ];

      walkthroughTriggers.forEach((trigger) => {
        it(`should invoke chat successfully from ${trigger}`, async () => {
          const args = [trigger];
          sandbox.stub(vscode.commands, "executeCommand").resolves();
          const res = await handlers.invokeTeamsAgent(args);
          chai.assert.isTrue(res.isOk());
          if (res.isOk()) {
            chai.assert.isTrue(res.value);
          } else {
            console.log(res.error);
          }
        });
      });
    });
  });

  describe("troubleshootSelectedText", async () => {
    beforeEach(() => {
      sandbox.stub(vscode.authentication, "getAccounts").resolves([
        {
          id: "someid",
          label: "",
        },
      ]);
    });
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
      sandbox
        .stub(vscode.extensions, "getExtension")
        .returns({ name: "github.copilot-chat" } as any);
      const error = new SystemError("test", "test", "test", "test");
      sandbox.stub(vscode.commands, "executeCommand").rejects(error);

      const res = await handlers.troubleshootSelectedText();
      chai.assert.isTrue(res.isErr());
    });
  });

  describe("troubleshootError", async () => {
    beforeEach(() => {
      sandbox.stub(vscode.authentication, "getAccounts").resolves([
        {
          id: "someid",
          label: "",
        },
      ]);
      sandbox.stub(globalState, "globalStateGet").resolves(true);
    });
    it("can invoke teams agent", async () => {
      sandbox
        .stub(vscode.extensions, "getExtension")
        .returns({ name: "github.copilot-chat" } as any);
      sandbox.stub(vscode.commands, "executeCommand").resolves();

      const currentError = new SystemError("test", "test", "test", "test");
      const res = await handlers.troubleshootError(["Notification", currentError]);
      chai.assert.isTrue(res.isOk());
    });

    it("missing args", async () => {
      const res = await handlers.troubleshootError([]);
      const calledCommand = sandbox.stub(vscode.commands, "executeCommand").resolves();
      chai.assert.isTrue(res.isOk());
      chai.assert.isFalse(calledCommand.calledOnce);
    });

    it("error", async () => {
      const error = new SystemError("test", "test", "test", "test");
      sandbox.stub(vscode.commands, "executeCommand").rejects(error);
      sandbox
        .stub(vscode.extensions, "getExtension")
        .returns({ name: "github.copilot-chat" } as any);

      const currentError = new SystemError("test", "test", "test", "test");
      const res = await handlers.troubleshootError(["triggerFrom", currentError]);
      chai.assert.isTrue(res.isErr());
    });
  });
});

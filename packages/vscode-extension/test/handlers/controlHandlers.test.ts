import { ok, TeamsAppManifest } from "@microsoft/teamsfx-api";
import { featureFlagManager, manifestUtils } from "@microsoft/teamsfx-core";
import * as projectSettingsHelper from "@microsoft/teamsfx-core/build/common/projectSettingsHelper";
import * as chai from "chai";
import fs from "fs-extra";
import * as sinon from "sinon";
import * as vscode from "vscode";
import { PanelType } from "../../src/controls/PanelType";
import { WebviewPanel } from "../../src/controls/webviewPanel";
import * as globalVariables from "../../src/globalVariables";
import {
  defaultWelcomePageKey,
  openFolderHandler,
  openLifecycleTreeview,
  openSamplesHandler,
  openWelcomeHandler,
  saveTextDocumentHandler,
  selectWalkthrough,
} from "../../src/handlers/controlHandlers";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetryTriggerFrom,
  TelemetryUpdateAppReason,
} from "../../src/telemetry/extTelemetryEvents";
import * as commonUtils from "../../src/utils/commonUtils";
import { getDefaultString } from "../../src/utils/localizeUtils";
import { getWalkThroughId } from "../../src/utils/projectStatusUtils";

describe("Control Handlers", () => {
  describe("openWelcomeHandler", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("opens normal walkthrough", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      sandbox.stub(manifestUtils, "readAppManifest").resolves(ok({} as TeamsAppManifest));
      sandbox.stub(manifestUtils, "getCapabilities").returns(["bot"]);
      const executeCommands = sandbox.stub(vscode.commands, "executeCommand");
      const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox
        .stub(vscode.window, "showQuickPick")
        .resolves({ label: getDefaultString("teamstoolkit.walkthroughs.title") });
      await openWelcomeHandler();

      sandbox.assert.calledOnceWithExactly(
        executeCommands,
        "workbench.action.openWalkthrough",
        "TeamsDevApp.ms-teams-vscode-extension#teamsToolkitGetStarted"
      );
    });

    it("opens default walkthrough", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      sandbox.stub(manifestUtils, "readAppManifest").resolves(ok({} as TeamsAppManifest));
      sandbox.stub(manifestUtils, "getCapabilities").returns(["bot"]);
      const executeCommands = sandbox.stub(vscode.commands, "executeCommand");
      const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      await openWelcomeHandler(TelemetryTriggerFrom.Auto, defaultWelcomePageKey);

      sandbox.assert.calledOnceWithExactly(
        executeCommands,
        "workbench.action.openWalkthrough",
        "TeamsDevApp.ms-teams-vscode-extension#teamsToolkitGetStarted"
      );
    });

    it("opens normal walkthrough -SideBar", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      sandbox.stub(manifestUtils, "readAppManifest").resolves(ok({} as TeamsAppManifest));
      sandbox.stub(manifestUtils, "getCapabilities").returns(["bot"]);
      const executeCommands = sandbox.stub(vscode.commands, "executeCommand");
      const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      await openWelcomeHandler("SideBar");
      sandbox.assert.calledOnceWithExactly(
        executeCommands,
        "workbench.action.openWalkthrough",
        "TeamsDevApp.ms-teams-vscode-extension#teamsToolkitGetStarted"
      );
    });

    it("opens walkthrough with chat", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      sandbox.stub(manifestUtils, "readAppManifest").resolves(ok({} as TeamsAppManifest));
      sandbox.stub(manifestUtils, "getCapabilities").returns(["bot"]);
      const executeCommands = sandbox.stub(vscode.commands, "executeCommand");
      const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox
        .stub(vscode.window, "showQuickPick")
        .resolves({ label: getDefaultString("teamstoolkit.walkthroughs.title") });
      await openWelcomeHandler();

      sandbox.assert.calledOnceWithExactly(
        executeCommands,
        "workbench.action.openWalkthrough",
        "TeamsDevApp.ms-teams-vscode-extension#teamsToolkitGetStartedWithChat"
      );
    });

    it("opens intelligent app walkthrough for API plugin apps", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      sandbox.stub(manifestUtils, "readAppManifest").resolves(ok({} as TeamsAppManifest));
      sandbox.stub(manifestUtils, "getCapabilities").returns(["plugin"]);
      sandbox.stub(globalVariables, "workspaceUri").value({ fsPath: "/test" });
      const executeCommands = sandbox.stub(vscode.commands, "executeCommand");
      const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      await openWelcomeHandler("invalidArgs");

      sandbox.assert.calledOnceWithExactly(
        executeCommands,
        "workbench.action.openWalkthrough",
        "TeamsDevApp.ms-teams-vscode-extension#buildIntelligentApps"
      );
    });

    it("opens intelligent app walkthrough with chat for API plugin apps", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      sandbox.stub(manifestUtils, "readAppManifest").resolves(ok({} as TeamsAppManifest));
      sandbox.stub(manifestUtils, "getCapabilities").returns(["plugin"]);
      sandbox.stub(globalVariables, "workspaceUri").value({ fsPath: "/test" });
      const executeCommands = sandbox.stub(vscode.commands, "executeCommand");
      const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      await openWelcomeHandler("invalidArgs");

      sandbox.assert.calledOnceWithExactly(
        executeCommands,
        "workbench.action.openWalkthrough",
        "TeamsDevApp.ms-teams-vscode-extension#buildIntelligentAppsWithChat"
      );
    });

    it("opens intelligent app walkthrough for JS/TS custom engine copilot apps", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      sandbox.stub(manifestUtils, "readAppManifest").resolves(ok({} as TeamsAppManifest));
      sandbox.stub(manifestUtils, "getCapabilities").returns(["bot"]);
      sandbox.stub(globalVariables, "workspaceUri").value({ fsPath: "/test" });
      sandbox.stub(fs, "pathExists").callsFake(async (path: string) => {
        return path.includes("package.json");
      });
      sandbox.stub(fs, "readFile").resolves(Buffer.from('"@microsoft/teams-ai"'));
      const executeCommands = sandbox.stub(vscode.commands, "executeCommand");
      const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      await openWelcomeHandler();

      sandbox.assert.calledOnceWithExactly(
        executeCommands,
        "workbench.action.openWalkthrough",
        "TeamsDevApp.ms-teams-vscode-extension#buildIntelligentApps"
      );
    });

    it("opens intelligent app walkthrough for python custom engine copilot apps", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      sandbox.stub(manifestUtils, "readAppManifest").resolves(ok({} as TeamsAppManifest));
      sandbox.stub(manifestUtils, "getCapabilities").returns(["bot"]);
      sandbox.stub(globalVariables, "workspaceUri").value({ fsPath: "/test" });
      sandbox.stub(fs, "pathExists").callsFake(async (path: string) => {
        return path.includes("requirements.txt");
      });
      sandbox.stub(fs, "readFile").resolves(Buffer.from("teams-ai"));
      const executeCommands = sandbox.stub(vscode.commands, "executeCommand");
      const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      await openWelcomeHandler();

      sandbox.assert.calledOnceWithExactly(
        executeCommands,
        "workbench.action.openWalkthrough",
        "TeamsDevApp.ms-teams-vscode-extension#buildIntelligentApps"
      );
    });

    it("opens normal walkthrough for JS/TS apps without ai library", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      sandbox.stub(manifestUtils, "readAppManifest").resolves(ok({} as TeamsAppManifest));
      sandbox.stub(manifestUtils, "getCapabilities").returns(["bot"]);
      sandbox.stub(globalVariables, "workspaceUri").value({ fsPath: "/test" });
      sandbox.stub(fs, "pathExists").callsFake(async (path: string) => {
        return path.includes("package.json");
      });
      sandbox.stub(fs, "readFile").resolves(Buffer.from(""));
      const executeCommands = sandbox.stub(vscode.commands, "executeCommand");
      const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox
        .stub(vscode.window, "showQuickPick")
        .resolves({ label: getDefaultString("teamstoolkit.walkthroughs.title") });
      await openWelcomeHandler();

      sandbox.assert.calledOnceWithExactly(
        executeCommands,
        "workbench.action.openWalkthrough",
        "TeamsDevApp.ms-teams-vscode-extension#teamsToolkitGetStarted"
      );
    });

    it("opens normal walkthrough for python custom engine copilot apps", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      sandbox.stub(manifestUtils, "readAppManifest").resolves(ok({} as TeamsAppManifest));
      sandbox.stub(manifestUtils, "getCapabilities").returns(["bot"]);
      sandbox.stub(globalVariables, "workspaceUri").value({ fsPath: "/test" });
      sandbox.stub(fs, "pathExists").callsFake(async (path: string) => {
        return path.includes("requirements.txt");
      });
      sandbox.stub(fs, "readFile").resolves(Buffer.from(""));
      const executeCommands = sandbox.stub(vscode.commands, "executeCommand");
      const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox
        .stub(vscode.window, "showQuickPick")
        .resolves({ label: getDefaultString("teamstoolkit.walkthroughs.title") });
      await openWelcomeHandler();

      sandbox.assert.calledOnceWithExactly(
        executeCommands,
        "workbench.action.openWalkthrough",
        "TeamsDevApp.ms-teams-vscode-extension#teamsToolkitGetStarted"
      );
    });
  });

  describe("openSamplesHandler", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("openSamplesHandler", async () => {
      const createOrShow = sandbox.stub(WebviewPanel, "createOrShow");
      const sendTelemetryEvent = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      await openSamplesHandler();

      sandbox.assert.calledOnceWithExactly(createOrShow, PanelType.SampleGallery, []);
    });
  });

  describe("openFolderHandler", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("empty args", async () => {
      const sendTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      const result = await openFolderHandler();

      chai.assert.isTrue(sendTelemetryStub.called);
      chai.assert.isTrue(result.isOk());
    });

    it("happy path", async () => {
      const sendTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      const openFolderInExplorerStub = sandbox.stub(commonUtils, "openFolderInExplorer");

      const result = await openFolderHandler("file://path/to/folder");

      chai.assert.isTrue(sendTelemetryStub.called);
      chai.assert.isTrue(openFolderInExplorerStub.calledOnceWith("/path/to/folder"));
      chai.assert.isTrue(result.isOk());
    });
  });

  describe("saveTextDocumentHandler", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("non valid project", () => {
      const isValidProjectStub = sandbox
        .stub(projectSettingsHelper, "isValidProject")
        .returns(false);
      sandbox.stub(globalVariables, "workspaceUri").value({ fsPath: "/path/to/workspace" });

      saveTextDocumentHandler({ document: {} } as any);

      chai.assert.isTrue(isValidProjectStub.calledOnceWith("/path/to/workspace"));
    });

    it("manual save reason", () => {
      const isValidProjectStub = sandbox
        .stub(projectSettingsHelper, "isValidProject")
        .returns(true);
      const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(globalVariables, "workspaceUri").value({ fsPath: "/path/to/workspace" });

      saveTextDocumentHandler({
        document: { fileName: "/dirname/fileName" },
        reason: vscode.TextDocumentSaveReason.Manual,
      } as vscode.TextDocumentWillSaveEvent);

      chai.assert.isTrue(isValidProjectStub.calledTwice);
      chai.assert.equal(isValidProjectStub.getCall(0).args[0], "/path/to/workspace");
      chai.assert.equal(isValidProjectStub.getCall(1).args[0], "/dirname");
      chai.assert.equal(sendTelemetryEventStub.getCall(0).args[0], TelemetryEvent.UpdateTeamsApp);
      chai.assert.deepEqual(sendTelemetryEventStub.getCall(0).args[1], {
        [TelemetryProperty.UpdateTeamsAppReason]: TelemetryUpdateAppReason.Manual,
      });
    });

    it("after delay save reason", () => {
      const isValidProjectStub = sandbox
        .stub(projectSettingsHelper, "isValidProject")
        .returns(true);
      const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(globalVariables, "workspaceUri").value({ fsPath: "/path/to/workspace" });

      saveTextDocumentHandler({
        document: { fileName: "/dirname/fileName" },
        reason: vscode.TextDocumentSaveReason.AfterDelay,
      } as vscode.TextDocumentWillSaveEvent);

      chai.assert.isTrue(isValidProjectStub.calledTwice);
      chai.assert.equal(isValidProjectStub.getCall(0).args[0], "/path/to/workspace");
      chai.assert.equal(isValidProjectStub.getCall(1).args[0], "/dirname");
      chai.assert.equal(sendTelemetryEventStub.getCall(0).args[0], TelemetryEvent.UpdateTeamsApp);
      chai.assert.deepEqual(sendTelemetryEventStub.getCall(0).args[1], {
        [TelemetryProperty.UpdateTeamsAppReason]: TelemetryUpdateAppReason.AfterDelay,
      });
    });

    it("focus out save reason", () => {
      const isValidProjectStub = sandbox
        .stub(projectSettingsHelper, "isValidProject")
        .callsFake((path: string | undefined) => {
          return path !== "/dirname";
        });
      const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(globalVariables, "workspaceUri").value({ fsPath: "/path/to/workspace" });

      saveTextDocumentHandler({
        document: { fileName: "/dirname/fileName" },
        reason: vscode.TextDocumentSaveReason.FocusOut,
      } as vscode.TextDocumentWillSaveEvent);

      chai.assert.isTrue(isValidProjectStub.calledThrice);
      chai.assert.equal(isValidProjectStub.getCall(0).args[0], "/path/to/workspace");
      chai.assert.equal(isValidProjectStub.getCall(1).args[0], "/dirname");
      chai.assert.equal(isValidProjectStub.getCall(2).args[0], "/");
      chai.assert.equal(sendTelemetryEventStub.getCall(0).args[0], TelemetryEvent.UpdateTeamsApp);
      chai.assert.deepEqual(sendTelemetryEventStub.getCall(0).args[1], {
        [TelemetryProperty.UpdateTeamsAppReason]: TelemetryUpdateAppReason.FocusOut,
      });
    });
  });

  describe("openLifecycleTreeview", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("TeamsFx Project", async () => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(globalVariables, "isTeamsFxProject").value(true);
      const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");

      await openLifecycleTreeview();

      chai.assert.isTrue(executeCommandStub.calledWith("teamsfx-lifecycle.focus"));
    });

    it("non-TeamsFx Project", async () => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(globalVariables, "isTeamsFxProject").value(false);
      const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");

      await openLifecycleTreeview();

      chai.assert.isTrue(executeCommandStub.calledWith("workbench.view.extension.teamsfx"));
    });
  });

  describe("selectWalkthrough", () => {
    let quickPickStub: sinon.SinonStub;
    let executeCommandStub: sinon.SinonStub;

    beforeEach(() => {
      // Stubbing VS Code APIs
      quickPickStub = sinon.stub(vscode.window, "showQuickPick");
      executeCommandStub = sinon.stub(vscode.commands, "executeCommand");
    });

    afterEach(() => {
      sinon.restore();
    });

    it("should select the Teams Toolkit walkthrough", async () => {
      quickPickStub.resolves({
        label: getDefaultString("teamstoolkit.walkthroughs.title"),
        detail: "Some description",
      });

      executeCommandStub.callsFake((command: string, ...args: any[]) => {
        chai.assert(command, "workbench.action.openWalkthrough");
        chai.assert(args[0], getWalkThroughId());
        return "Success";
      });

      const result = await selectWalkthrough();

      chai.assert.isTrue(quickPickStub.calledOnce);
      chai.assert.isTrue(executeCommandStub.calledOnce);
      chai.assert.isTrue(result.isOk());
    });

    it("should select the Teams Toolkit walkthrough with chat", async () => {
      sinon.stub(featureFlagManager, "getBooleanValue").returns(true);
      quickPickStub.callsFake(
        (items: vscode.QuickPickItem[], options?: vscode.QuickPickOptions) => {
          chai.assert.isTrue(items.length == 2);
          chai.assert.isTrue(
            items[0].detail === getDefaultString("teamstoolkit.walkthroughs.withChat.description")
          );
          return Promise.resolve({
            label: getDefaultString("teamstoolkit.walkthroughs.title"),
            detail: "Some description",
          });
        }
      );

      executeCommandStub.callsFake((command: string, ...args: any[]) => {
        chai.assert(command, "workbench.action.openWalkthrough");
        chai.assert(args[0], getWalkThroughId());
        return "Success";
      });

      const result = await selectWalkthrough();

      chai.assert.isTrue(quickPickStub.calledOnce);
      chai.assert.isTrue(executeCommandStub.calledOnce);
      chai.assert.isTrue(result.isOk());
    });

    it("should select the declarative agent walkthrough", async () => {
      quickPickStub.resolves({
        label: getDefaultString("teamstoolkit.walkthroughs.buildIntelligentApps.title"),
        detail: "Some description",
      });

      executeCommandStub.callsFake((command: string, ...args: any[]) => {
        chai.assert(command, "workbench.action.openWalkthrough");
        chai.assert(args[0], "TeamsDevApp.ms-teams-vscode-extension#buildIntelligentApps");
        return "Success";
      });

      const result = await selectWalkthrough();

      chai.assert.isTrue(quickPickStub.calledOnce);
      chai.assert.isTrue(executeCommandStub.calledOnce);
      chai.assert.isTrue(result.isOk());
    });

    it("should select the declarative agent walkthrough - empty label", async () => {
      quickPickStub.resolves({
        detail: "Some description",
      });

      executeCommandStub.callsFake((command: string, ...args: any[]) => {
        chai.assert(command, "workbench.action.openWalkthrough");
        chai.assert(args[0], "TeamsDevApp.ms-teams-vscode-extension#buildIntelligentApps");
        return "Success";
      });

      const result = await selectWalkthrough();

      chai.assert.isTrue(quickPickStub.calledOnce);
      chai.assert.isTrue(executeCommandStub.calledOnce);
      chai.assert.isTrue(result.isOk());
    });
  });
});
